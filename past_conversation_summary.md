# Summary: Merging Shopify Reports Data (April 29, 2026)

This document summarizes the development of the automated sales reporting system which consolidates data from Shopify, Shopee, and Lazada into a master Excel report.

---

## 🎯 Primary Objective
To automate the creation of the `Full_Sales_Report.xlsx` based on the formatting requirements of `Sample Reports.xlsx`. The system needed to handle different data structures from various platforms and ensure accurate financial mapping.

---

## 🛠️ Implementation Details

### 1. Data Sources
*   **Shopify**: Parsed from `orders_export(1).csv` and `transactions_export.csv`.
*   **Shopee**: Parsed from `Shopee Transaction Report.xlsx` (specifically the `orders` sheet).
*   **Lazada**: Parsed from `Lazada Transaction Report.xlsx`.
*   **Master Format**: `Sample Reports.xlsx` served as the template for styling, headers, and sheet structure.

### 2. Key Logic & Business Rules
*   **PayMongo Fee Calculation**: We determined from sample data that the PayMongo fee rate is **1.5%**. The script automatically calculates this for Shopify orders using that gateway.
*   **Shopee Fees**: Shopee's "Service Fee" is automatically mapped to the "Less payment fee" column in the summary.
*   **Lazada Grouping**: Lazada orders are grouped by `orderNumber` to ensure multi-item orders are handled correctly as a single transaction in the summary.
*   **Status Handling**: Logic was implemented to identify "Paid", "Pending", and "Cancelled" statuses across all platforms to ensure only valid sales are totaled.

### 3. SKU Normalization & Offtake Reports
One of the most complex parts was ensuring that product names from all three platforms matched the specific SKU names in your master report.
*   **Normalization Map**: I built a mapping system that translates various product names (e.g., "Fresh Bamboo 500", "Handwash") into the exact names used in the **Offtake Report Summary per sku** sheet.
*   **Automated Offtake**: The script tallies quantities sold per SKU per month across all platforms and populates the report automatically.
*   **Formula Preservation**: (Updated May 6) The script now copies entire sheet structures including **Excel formulas** (e.g., `% contribution`, `vs PY`), ensuring the summary sheet is fully dynamic.
*   **Date Matching**: (Updated May 6) Improved logic to map 2026 sales data accurately to the template's calendar months.
*   **Channel Preservation**: (Updated May 14) The script copies all channel rows from the template, including physical store channels, maintaining the full layout of the Offtake Report Summary.

---

## 📄 Output Structure: `Full_Sales_Report.xlsx`
The generated report includes four automated sheets:
1.  **Daily Sales Summary**: A merged chronological list of all orders.
2.  **Real Time Inventory Report**: A template copy ready for manual inventory input.
3.  **Offtake Report Summary**: Automatic monthly totals per sales channel (Shopify/Shopee/Lazada).
4.  **Offtake Report Summary per sku**: SKU-level quantity tracking by month.

---

## 🚀 How to Run
You can regenerate this report at any time by running the following command in your terminal:
```bash
python3 generate_report.py
```

> [!NOTE]
> The script will automatically scan the current directory for the latest CSV and XLSX export files from Shopify, Shopee, and Lazada.
