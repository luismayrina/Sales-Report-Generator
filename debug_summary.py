import json
import os
import re
from sales_report_app.core.physical_parser import load_physical_channels
from sales_report_app.core.report_generator import generate_full_report

offtake_path = "docs/2026 OFFTAKE REPORT.xlsx"
orders = load_physical_channels(offtake_path, current_year=2026)
print(f"Total physical orders loaded: {len(orders)}")

# Check what the actual 'source' and 'net' are
cr_sales = [o for o in orders if o.get('source') == 'Common Room' and o.get('date') and o.get('date').month == 1]
total = sum(o.get('net', 0) for o in cr_sales)
print(f"Total calculated for Common Room Jan: {total}")

