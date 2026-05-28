import sys
import argparse
import os

def run_cli():
    print("Running in CLI mode...")
    from core.shopify_parser import process_shopify
    from core.shopee_parser import load_shopee
    from core.lazada_parser import load_lazada
    from core.physical_parser import load_physical_channels, load_2025_data
    from core.report_generator import generate_full_report
    
    base_dir = os.getcwd()
    
    # Default file paths
    orders_csv = os.path.join(base_dir, "orders_export(1).csv")
    txns_csv = os.path.join(base_dir, "transactions_export.csv")
    shopee_xlsx = os.path.join(base_dir, "Shopee Transaction Report.xlsx")
    lazada_xlsx = os.path.join(base_dir, "Lazada Transaction Report.xlsx")
    template_xlsx = os.path.join(base_dir, "Sample Reports.xlsx")
    output_xlsx = os.path.join(base_dir, "Full_Sales_Report.xlsx")
    
    offtake_xlsx = os.path.join(base_dir, "docs", "2026 OFFTAKE REPORT.xlsx")
    if not os.path.exists(offtake_xlsx):
        offtake_xlsx = os.path.join(base_dir, "2026 OFFTAKE REPORT.xlsx")
    
    missing = []
    for path, name in [
        (orders_csv, "Shopify Orders"),
        (txns_csv, "Shopify Transactions"),
        (shopee_xlsx, "Shopee Report"),
        (lazada_xlsx, "Lazada Report"),
        (template_xlsx, "Template"),
        (offtake_xlsx, "Offtake Report (2026 OFFTAKE REPORT.xlsx)")
    ]:
        if not os.path.exists(path):
            missing.append(name)
            
    if missing:
        print(f"Error: Missing the following required files in the current directory: {', '.join(missing)}")
        sys.exit(1)
        
    print("📂  Loading Shopify transactions & orders…")
    shopify_orders = process_shopify(orders_csv, txns_csv)
    
    print("📂  Loading Shopee transactions…")
    shopee_orders = load_shopee(shopee_xlsx)
    
    print("📂  Loading Lazada transactions…")
    lazada_orders = load_lazada(lazada_xlsx)
    
    print("📂  Loading Offtake Report (Physical channels & TikTok)…")
    physical_orders = load_physical_channels(offtake_xlsx)
    py_data = load_2025_data(offtake_xlsx)
    
    all_orders = shopify_orders + shopee_orders + lazada_orders + physical_orders
    print(f"\n📊  Total combined orders: {len(all_orders)}")
    
    def cli_logger(msg):
        print(msg)
        
    success, result_path = generate_full_report(
        all_orders=all_orders,
        template_path=template_xlsx,
        output_path=output_xlsx,
        py_data=py_data,
        logger=cli_logger
    )
    
    if success:
        print(f"\n✅ Report generated successfully at: {result_path}")
    else:
        print(f"\n❌ Failed to generate report: {result_path}")

def main():
    parser = argparse.ArgumentParser(description="Sales Report Generator Desktop App")
    parser.add_argument("--cli", action="store_true", help="Run in Command Line mode (no GUI)")
    args = parser.parse_args()

    if args.cli:
        run_cli()
    else:
        try:
            from PySide6.QtWidgets import QApplication
            from ui.main_window import MainWindow
            
            app = QApplication(sys.argv)
            window = MainWindow()
            window.show()
            sys.exit(app.exec())
        except ImportError as e:
            print(f"Error starting GUI: {e}")
            print("Please ensure PySide6 is installed. You can also run with --cli for console mode.")
            sys.exit(1)

if __name__ == "__main__":
    main()
