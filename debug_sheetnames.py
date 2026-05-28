import openpyxl
from sales_report_app.core.physical_parser import SHEET_CHANNEL_MAP

wb = openpyxl.load_workbook("docs/2026 OFFTAKE REPORT.xlsx", read_only=True)
for sheetname in wb.sheetnames:
    print(f"Sheet: '{sheetname}'")
    channel_name = SHEET_CHANNEL_MAP.get(sheetname)
    if not channel_name:
        channel_name = sheetname.title()
    print(f"  -> Maps to: '{channel_name}'")

