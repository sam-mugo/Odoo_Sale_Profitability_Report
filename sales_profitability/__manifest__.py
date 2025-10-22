{
    'name': 'Sales Profitability Report',
    'version': '18.0.1.0.0',
    'category': 'Sales',
    'summary': 'Advanced sales profitability analysis with order-wise revenue, cost, and margin reporting',
    'description': """
        Sales Profitability Report Module
        =================================
        
        Features:
        - Order-wise revenue, cost, and margin analysis
        - Advanced filtering by date range, product category, and customer
        - Excel export functionality
        - Professional QWeb report printing
        - Interactive wizard interface
        - Real-time profitability calculations
    """,
    'author': 'Your Company',
    'website': 'https://www.yourcompany.com',
    'depends': ['base', 'sale', 'stock', 'product'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/sales_profitability_wizard_view.xml',
        'report/sales_profitability_report.xml',
        'report/sales_profitability_template.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'sales_profitability/static/src/css/profitability_report.css',
        ],
    },
    'installable': True,
    'auto_install': True,
    'application': True,
    'license': 'LGPL-3',
}