from sales_report_app.core.physical_parser import load_physical_channels
import os

offtake_path = "docs/2026 OFFTAKE REPORT.xlsx"
orders = load_physical_channels(offtake_path, current_year=2026)

sources = set(o.get('source') for o in orders)
print("Sources:", sources)

cr_sales = [o for o in orders if o.get('source') == 'Common Room' and o.get('date') and o.get('date').month == 1]
total = sum(o.get('net', 0) for o in cr_sales)
print(f"Total calculated for Common Room Jan: {total}")
