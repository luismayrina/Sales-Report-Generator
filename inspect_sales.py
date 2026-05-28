from sales_report_app.core.physical_parser import load_physical_channels
import os

offtake_path = "docs/2026 OFFTAKE REPORT.xlsx"
orders = load_physical_channels(offtake_path, current_year=2026, py_year="2025")

common_room_jan_sales = 0
common_room_feb_sales = 0

for o in orders:
    if o.get("source", "").upper() == "COMMON ROOM":
        dt = o.get("date")
        net = o.get("net", 0)
        if dt:
            if dt.month == 1:
                common_room_jan_sales += net
            elif dt.month == 2:
                common_room_feb_sales += net

print(f"Common Room Jan Sales: {common_room_jan_sales}")
print(f"Common Room Feb Sales: {common_room_feb_sales}")
