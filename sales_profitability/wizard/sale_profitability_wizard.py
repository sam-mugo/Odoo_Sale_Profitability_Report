from odoo import models, fields, api, _
from odoo.exceptions import UserError
import xlsxwriter
import base64
from io import BytesIO
from datetime import datetime, timedelta


class SalesProfitabilityWizard(models.TransientModel):
    _name = 'sales.profitability.wizard'
    _description = 'Sales Profitability Report Wizard'

    # Filter Fields
    date_from = fields.Date(
        string='Date From', 
        required=True, 
        default=lambda self: fields.Date.today().replace(day=1)
    )
    date_to = fields.Date(
        string='Date To', 
        required=True, 
        default=fields.Date.today
    )
    partner_ids = fields.Many2many(
        'res.partner', 
        string='Customers',
        domain=[('is_company', '=', True), ('customer_rank', '>', 0)]
    )
    category_ids = fields.Many2many(
        'product.category', 
        string='Product Categories'
    )
    company_id = fields.Many2one(
        'res.company', 
        string='Company', 
        default=lambda self: self.env.company
    )
    
    # Report Options
    group_by = fields.Selection([
        ('order', 'By Sales Order'),
        ('customer', 'By Customer'),
        ('category', 'By Product Category'),
        ('product', 'By Product')
    ], string='Group By', default='order', required=True)
    
    show_details = fields.Boolean(string='Show Line Details', default=True)
    include_taxes = fields.Boolean(string='Include Taxes in Revenue', default=True)
    
    # Results
    report_data = fields.Text(string='Report Data')
    excel_file = fields.Binary(string='Excel File')
    excel_filename = fields.Char(string='Excel Filename')

    @api.model
    def _get_domain_filters(self):
        """Build domain filters based on wizard selections"""
        domain = [
            ('state', 'in', ['sale', 'done']),
            ('date_order', '>=', self.date_from),
            ('date_order', '<=', self.date_to),
            ('company_id', '=', self.company_id.id)
        ]
        
        if self.partner_ids:
            domain.append(('partner_id', 'in', self.partner_ids.ids))
            
        return domain

    @api.model
    def _calculate_order_costs(self, order_line):
        """Calculate total cost for an order line including all components"""
        total_cost = 0.0
        
        # Get product standard cost
        product_cost = order_line.product_id.standard_price
        total_cost += product_cost * order_line.product_uom_qty
        
        # Add landed costs if available
        if hasattr(order_line, 'landed_cost_value'):
            total_cost += order_line.landed_cost_value or 0.0
            
        # Add additional costs from stock moves if order is delivered
        if order_line.move_ids:
            for move in order_line.move_ids.filtered(lambda m: m.state == 'done'):
                if move.stock_valuation_layer_ids:
                    total_cost += sum(move.stock_valuation_layer_ids.mapped('value'))
                    
        return total_cost

    @api.model
    def _get_profitability_data(self):
        """Main method to calculate profitability data"""
        domain = self._get_domain_filters()
        orders = self.env['sale.order'].search(domain)
        
        if not orders:
            raise UserError(_('No sales orders found for the selected criteria.'))
        
        profitability_data = []
        
        for order in orders:
            order_data = {
                'order_id': order.id,
                'order_name': order.name,
                'customer': order.partner_id.name,
                'customer_id': order.partner_id.id,
                'date_order': order.date_order,
                'currency': order.currency_id.name,
                'lines': [],
                'totals': {
                    'revenue': 0.0,
                    'cost': 0.0,
                    'margin': 0.0,
                    'margin_percent': 0.0
                }
            }
            
            # Process each order line
            for line in order.order_line:
                # Skip delivery and discount lines
                if line.is_delivery or line.product_id.type == 'service':
                    continue
                    
                # Filter by product category if specified
                if self.category_ids and line.product_id.categ_id not in self.category_ids:
                    continue
                
                # Calculate revenue
                if self.include_taxes:
                    line_revenue = line.price_total
                else:
                    line_revenue = line.price_subtotal
                
                # Calculate cost
                line_cost = self._calculate_order_costs(line)
                
                # Calculate margin
                line_margin = line_revenue - line_cost
                line_margin_percent = (line_margin / line_revenue * 100) if line_revenue else 0.0
                
                line_data = {
                    'product_name': line.product_id.name,
                    'product_code': line.product_id.default_code or '',
                    'category': line.product_id.categ_id.name,
                    'quantity': line.product_uom_qty,
                    'unit_price': line.price_unit,
                    'revenue': line_revenue,
                    'cost': line_cost,
                    'margin': line_margin,
                    'margin_percent': line_margin_percent,
                    'uom': line.product_uom.name
                }
                
                order_data['lines'].append(line_data)
                
                # Update order totals
                order_data['totals']['revenue'] += line_revenue
                order_data['totals']['cost'] += line_cost
                order_data['totals']['margin'] += line_margin
            
            # Calculate overall margin percentage
            if order_data['totals']['revenue']:
                order_data['totals']['margin_percent'] = (
                    order_data['totals']['margin'] / order_data['totals']['revenue'] * 100
                )
            
            # Only include orders with lines (after filtering)
            if order_data['lines']:
                profitability_data.append(order_data)
        
        return profitability_data

    def action_generate_report(self):
        """Generate and display the profitability report"""
        try:
            report_data = self._get_profitability_data()
            
            # Store report data for the report template
            self.report_data = str(report_data)
            
            # Return action to display the report
            return {
                'type': 'ir.actions.report',
                'report_name': 'sales_profitability.profitability_report',
                'report_type': 'qweb-html',
                'data': {'report_data': report_data, 'wizard_id': self.id},
                'context': self.env.context,
            }
            
        except Exception as e:
            raise UserError(_('Error generating report: %s') % str(e))

    def action_export_excel(self):
        """Export profitability data to Excel"""
        try:
            report_data = self._get_profitability_data()
            
            # Create Excel file
            output = BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True, 'font_size': 12, 'bg_color': '#D3D3D3',
                'border': 1, 'align': 'center', 'valign': 'vcenter'
            })
            
            subheader_format = workbook.add_format({
                'bold': True, 'font_size': 10, 'bg_color': '#F0F0F0',
                'border': 1, 'align': 'left'
            })
            
            data_format = workbook.add_format({'border': 1, 'align': 'left'})
            number_format = workbook.add_format({'border': 1, 'num_format': '#,##0.00'})
            percent_format = workbook.add_format({'border': 1, 'num_format': '0.00%'})
            
            # Create Summary sheet
            summary_sheet = workbook.add_worksheet('Profitability Summary')
            
            # Write headers
            headers = [
                'Order #', 'Customer', 'Date', 'Revenue', 'Cost', 
                'Margin', 'Margin %', 'Currency'
            ]
            
            for col, header in enumerate(headers):
                summary_sheet.write(0, col, header, header_format)
            
            # Write data
            row = 1
            total_revenue = total_cost = total_margin = 0.0
            
            for order_data in report_data:
                summary_sheet.write(row, 0, order_data['order_name'], data_format)
                summary_sheet.write(row, 1, order_data['customer'], data_format)
                summary_sheet.write(row, 2, order_data['date_order'].strftime('%Y-%m-%d'), data_format)
                summary_sheet.write(row, 3, order_data['totals']['revenue'], number_format)
                summary_sheet.write(row, 4, order_data['totals']['cost'], number_format)
                summary_sheet.write(row, 5, order_data['totals']['margin'], number_format)
                summary_sheet.write(row, 6, order_data['totals']['margin_percent'] / 100, percent_format)
                summary_sheet.write(row, 7, order_data['currency'], data_format)
                
                total_revenue += order_data['totals']['revenue']
                total_cost += order_data['totals']['cost']
                total_margin += order_data['totals']['margin']
                
                row += 1
            
            # Write totals
            row += 1
            summary_sheet.write(row, 0, 'TOTAL', header_format)
            summary_sheet.write(row, 3, total_revenue, number_format)
            summary_sheet.write(row, 4, total_cost, number_format)
            summary_sheet.write(row, 5, total_margin, number_format)
            if total_revenue:
                summary_sheet.write(row, 6, (total_margin / total_revenue), percent_format)
            
            # Auto-adjust column widths
            summary_sheet.set_column('A:A', 15)
            summary_sheet.set_column('B:B', 25)
            summary_sheet.set_column('C:C', 12)
            summary_sheet.set_column('D:G', 15)
            summary_sheet.set_column('H:H', 10)
            
            # Create Detailed sheet if requested
            if self.show_details:
                detail_sheet = workbook.add_worksheet('Order Line Details')
                
                detail_headers = [
                    'Order #', 'Customer', 'Product Code', 'Product Name',
                    'Category', 'Quantity', 'UoM', 'Unit Price',
                    'Revenue', 'Cost', 'Margin', 'Margin %'
                ]
                
                for col, header in enumerate(detail_headers):
                    detail_sheet.write(0, col, header, header_format)
                
                detail_row = 1
                for order_data in report_data:
                    for line in order_data['lines']:
                        detail_sheet.write(detail_row, 0, order_data['order_name'], data_format)
                        detail_sheet.write(detail_row, 1, order_data['customer'], data_format)
                        detail_sheet.write(detail_row, 2, line['product_code'], data_format)
                        detail_sheet.write(detail_row, 3, line['product_name'], data_format)
                        detail_sheet.write(detail_row, 4, line['category'], data_format)
                        detail_sheet.write(detail_row, 5, line['quantity'], number_format)
                        detail_sheet.write(detail_row, 6, line['uom'], data_format)
                        detail_sheet.write(detail_row, 7, line['unit_price'], number_format)
                        detail_sheet.write(detail_row, 8, line['revenue'], number_format)
                        detail_sheet.write(detail_row, 9, line['cost'], number_format)
                        detail_sheet.write(detail_row, 10, line['margin'], number_format)
                        detail_sheet.write(detail_row, 11, line['margin_percent'] / 100, percent_format)
                        detail_row += 1
                
                # Auto-adjust column widths for detail sheet
                detail_sheet.set_column('A:B', 15)
                detail_sheet.set_column('C:C', 12)
                detail_sheet.set_column('D:D', 30)
                detail_sheet.set_column('E:E', 20)
                detail_sheet.set_column('F:L', 12)
            
            workbook.close()
            output.seek(0)
            
            # Create filename
            filename = f'Sales_Profitability_{self.date_from}_{self.date_to}.xlsx'
            
            # Save file
            self.excel_file = base64.b64encode(output.read())
            self.excel_filename = filename
            
            # Return download action
            return {
                'type': 'ir.actions.act_url',
                'url': f'/web/content/?model=sales.profitability.wizard&id={self.id}&field=excel_file&download=true&filename={filename}',
                'target': 'self',
            }
            
        except Exception as e:
            raise UserError(_('Error creating Excel file: %s') % str(e))

    def action_print_report(self):
        """Print the profitability report"""
        try:
            report_data = self._get_profitability_data()
            
            return {
                'type': 'ir.actions.report',
                'report_name': 'sales_profitability.profitability_report',
                'report_type': 'qweb-pdf',
                'data': {'report_data': report_data, 'wizard_id': self.id},
                'context': self.env.context,
            }
            
        except Exception as e:
            raise UserError(_('Error printing report: %s') % str(e))