from sales_report_app.core.physical_parser import load_physical_channels, load_py_data
import os

offtake_path = "docs/2026 OFFTAKE REPORT.xlsx"
orders = load_physical_channels(offtake_path, current_year=2026)
py_data = load_py_data(offtake_path, py_year=2025)

common_room_jan_sales = 0
for o in orders:
    if o.get("source", "").upper() == "COMMON ROOM":
        dt = o.get("date")
        net = o.get("net", 0)
        if dt and dt.month == 1:
            common_room_jan_sales += net

print(f"Common Room Jan 2026 Actual: {common_room_jan_sales}")
print(f"Common Room Jan 2025 PY: {py_data.get('Common Room', {}).get(1)}")
