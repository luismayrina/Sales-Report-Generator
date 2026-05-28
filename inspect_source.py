from sales_report_app.core.physical_parser import load_physical_channels

offtake_path = "docs/2026 OFFTAKE REPORT.xlsx"
orders = load_physical_channels(offtake_path, current_year=2026)

sources = set(o.get('source') for o in orders)
print(sources)
