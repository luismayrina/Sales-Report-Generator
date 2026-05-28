from sales_report_app.core.physical_parser import load_py_data

py_data = load_py_data("docs/2026 OFFTAKE REPORT.xlsx", py_year="2025")
print(py_data.keys())
