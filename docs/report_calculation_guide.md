# Sales Report Calculation Guide

## 1. Overview
The Sales Report Generator takes sales and transaction exports from Shopify, Shopee, and Lazada, along with a Sample Reports Template, and consolidates them into a unified, clean Excel workbook (`Full_Sales_Report.xlsx`). 

It automates the process of merging financial data, matching platform-specific SKUs to normalized product names, grouping multi-item orders, and generating high-level monthly summaries for sales channels and individual inventory items.

## 2. Input Files

### Shopify Orders CSV
- **Expected File**: CSV format.
- **Required Columns**: `Name`, `Financial Status`, `Paid at`, `Payment Method`, `Total`, `Discount Amount`, `Shipping`, `Subtotal`, `Created at`, `Billing Name`, `Lineitem name`, `Lineitem quantity`, `Lineitem price`, `Payment Reference`.
- **Purpose**: Provides order-level totals, line items, and customer information.
- **Used In**: Daily Sales Summary, Offtake Report Summary, Offtake Report Summary per sku.

### Shopify Transactions CSV
- **Expected File**: CSV format.
- **Required Columns**: `Name`, `Status`, `Gateway`, `Amount`, `Created At`.
- **Purpose**: Provides accurate payment gateway and deposit date information, which is critical for determining PayMongo fees.
- **Used In**: Daily Sales Summary.

### Shopee Transaction Report XLSX
- **Expected File**: Excel format (XLSX).
- **Required Sheet**: First active sheet (default sheet).
- **Required Columns**: `Order ID`, `Order Status`, `Order Paid Time`, `Order Creation Date`, `Total Buyer Payment`, `Service Fee`, `Total Discount(PHP)`, `Buyer Paid Shipping Fee`, `Grand Total`, `Product Name`, `Quantity`, `Original Price`, `Receiver Name`.
- **Purpose**: Provides Shopee sales, fees, and ordered items.
- **Used In**: Daily Sales Summary, Offtake Report Summary, Offtake Report Summary per sku.

### Lazada Transaction Report XLSX
- **Expected File**: Excel format (XLSX).
- **Required Sheet**: First active sheet (default sheet).
- **Required Columns**: `orderNumber`, `status`, `createTime`, `deliveredDate`, `paidPrice`, `unitPrice`, `sellerDiscountTotal`, `shippingFee`, `itemName`, `customerName`.
- **Purpose**: Provides Lazada sales and grouped line items.
- **Used In**: Daily Sales Summary, Offtake Report Summary, Offtake Report Summary per sku.

### Sample Reports Template XLSX
- **Expected File**: Excel format (XLSX).
- **Required Sheets**: 
  - `Daily Sales Summary ` (Note the trailing space)
  - `Real Time Inventory Report`
  - `Offtake Report Summary`
  - `Offtake Report Summary per sku`
- **Purpose**: Provides the visual formatting, headers, existing formulas, and baseline structure for the output report.
- **Used In**: All output sheets.

---

## 3. Sheet-by-Sheet Calculation Details

### Sheet 1: Daily Sales Summary

- **Purpose**: A comprehensive, chronological ledger of all orders across all three platforms.
- **Source Files**: Shopify Orders, Shopify Transactions, Shopee Report, Lazada Report.
- **Sorting**: Orders are sorted chronologically by the creation date. If dates match, the order is: Shopify -> Shopee -> Lazada.
- **Exclusions**: Shopify orders with empty `Financial Status` are excluded.

| Output Column | Source Platform | Source Column / Logic | Calculation Rule | Notes |
|---|---|---|---|---|
| A (source) | All | Internal | "Shopify", "Shopee", or "Lazada" | |
| B (Date) | All | Created at / Order Creation Date / createTime | Parsed as date | Formatted as `d-mmm` |
| C (Customer Name) | All | Billing Name / Receiver Name / customerName | Direct text copy | |
| D (Order No.) | All | Name / Order ID / orderNumber | Direct text copy | |
| E (Gross Amount) | Shopify | Subtotal + Discount Amount | Gross before discounts | |
| E (Gross Amount) | Shopee | Total Buyer Payment | Sum across all grouped rows | |
| E (Gross Amount) | Lazada | unitPrice | Sum across all grouped rows | |
| F (Plus Shipping fee) | All | Shipping / Buyer Paid Shipping Fee / shippingFee | Sum across grouped rows (if any) | |
| G (Less payment fee) | Shopee | Service Fee | Sum across grouped rows | Shopify/Lazada is blank here |
| I (Less Discount) | All | Discount Amount / Total Discount(PHP) / sellerDiscountTotal | Sum across grouped rows | Absolute value for Lazada |
| R (Less Paymongo fee) | Shopify | Gateway + Total | 1.5% of Total if gateway is PayMongo | Only applies to PayMongo |
| T (Net Amount) | Shopify | Total | Direct text copy | |
| T (Net Amount) | Shopee | Grand Total | Taken from the *first* row of the order group | |
| T (Net Amount) | Lazada | paidPrice | Sum across all grouped rows | |
| U (Payout/Date Deposit) | All | Paid at / Order Paid Time / deliveredDate | Only populated if Paid/Delivered | |
| V (BANK) | Shopify | Payment Method / Gateway / Status | "UBP" if PayMongo, "BANK DEPOSIT" or "MANUAL" | "PENDING" if not paid |
| V (BANK) | Shopee | Status | "UBP" if paid, "CANCELLED" if cancelled | |
| V (BANK) | Lazada | Status | "LAZADA" if paid, "CANCELLED" if cancelled | |
| X (Date Deposit) | All | Same as Column U | Copied directly | |
| Y (Reference No.) | Shopify | Payment Reference | Direct text copy | Only applies to Shopify |

#### Platform-Specific Rules:
- **Shopify**: Orders and Transactions are merged using the order `Name`. PayMongo detection looks for the word "paymongo" in either the Gateway or Payment Method.
- **Shopee**: Rows sharing the same `Order ID` are grouped into a single transaction. Item-level values (Service Fee, Discounts) are summed, while the `Grand Total` is taken entirely from the first row of the group.
- **Lazada**: Rows sharing the same `orderNumber` are grouped. Order status relies on "canceled" (is_cancelled) or "delivered/confirmed" (paid). If delivered, the `deliveredDate` becomes the deposit date.

---

### Sheet 2: Real Time Inventory Report

- **Purpose**: Provides a snapshot of inventory.
- **Population**: This sheet is a direct visual copy from the `Sample Reports Template.xlsx`. 
- **Notes**: The script **does not** automatically calculate or write inventory levels. All existing rows, styling, and formulas are duplicated from the template.
- **Pending Implementation**: Cell `A3` is overwritten with the text `[Pending Data Source for Finished Goods + Raw Materials]` to indicate where future data should be injected.

| Area / Column | Source | Calculation Rule | Notes |
|---|---|---|---|
| All Columns | Template | Copied exactly | Preserves styles and formulas |
| Cell A3 | Script | Hardcoded text | Placeholder for future raw material logic |

---

### Sheet 3: Offtake Report Summary

- **Purpose**: A high-level monthly revenue breakdown by sales channel.
- **Calculation**: Monthly totals are calculated by summing the `Net Amount` of all non-cancelled orders that fall within a specific month.
- **Overwrites**: The script finds the exact row for Shopify, Shopee, and Lazada, and overwrites the monthly columns.
- **Preservation**: Physical store rows, `% contribution` formulas, and `vs PY` columns are preserved identically from the template.

| Row / Channel | Source Platform | Month Mapping Logic | Value Written | Notes |
|---|---|---|---|---|
| Shopify | Shopify | Order Date month | Sum of Net Amount | Excludes cancelled orders |
| Shopee | Shopee | Order Date month | Sum of Net Amount | Excludes cancelled orders |
| Lazada | Lazada | Order Date month | Sum of Net Amount | Excludes cancelled orders |
| Physical Store Rows | Template | Preserved | Not overwritten | Copied from template |

---

### Sheet 4: Offtake Report Summary per sku

- **Purpose**: Tracks exactly how many units of a specific SKU were sold each month.
- **Mechanism**: The script reads the template to discover all "valid SKUs" listed in the `VARIANT` column. When processing orders, the raw product name is normalized and matched against this list.
- **Grouping**: Quantities are grouped by the matched SKU string and the month of the order date. Cancelled orders are ignored.

| Output SKU Name | Possible Source Product Names | Source Platforms | Quantity Logic | Notes |
|---|---|---|---|---|
| White Tea 500ml | "white tea 500" | All | Sum of Lineitem quantity | Matched via normalization |
| Fresh Bamboo | "fresh bamboo" | All | Sum of Lineitem quantity | Matched via normalization |

**SKU Normalization Rules:**
Product names are converted to lowercase. The `normalize_sku` function maps substrings (e.g., "relaxing naturals 500") to official names (e.g., "Relaxing Naturals 500ml"). The normalized string is then checked to see if it exists as a substring within any of the valid SKUs extracted from the template.

---

### Sheet 5: Unmapped SKUs

- **Purpose**: To catch and log products that the script failed to identify.
- **Condition**: Created **only** if there are items in the orders that do not map to any of the valid SKUs listed in the Offtake Summary template.
- **User Action**: The user must review this sheet. If a product is legitimate, the template's `VARIANT` column should be updated, or the developer must update the `mapping` dictionary in `core/sku_normalizer.py`.

**Expected Columns:**
- Platform
- Order No
- Date
- Raw Name
- Normalized
- Qty

---

## 4. Business Rules

### PayMongo Fee
- **Rate**: 1.5% (`0.015`).
- **Platform**: Shopify only.
- **Trigger**: "paymongo" must appear in the Gateway or Payment Method column.
- **Formula**: `round(Total * 0.015, 2)`
- **Example**: Shopify order Total is ₱1,000, Gateway is "PayMongo". Fee = `1000 * 0.015 = 15.00`.

### Shopee Fees
- **Rule**: The script sums the `Service Fee` across all line items of a grouped order and maps it directly to the `Less payment fee` column in the Daily Sales Summary.
- **Impact**: It does not alter the Net Sales, as Shopee's `Grand Total` is used directly for Net.

### Lazada Grouping
- **Rule**: Lazada exports contain one row per item. The script groups them using `orderNumber`.
- **Sums**: `paidPrice`, `unitPrice`, `sellerDiscountTotal`, and `shippingFee` are summed across the group.
- **Discounts**: Negative discounts are converted to positive absolute values using `abs()`.

### Status Handling
- **Shopify**: Determines paid status via `Financial Status` == "paid".
- **Shopee**: Determines paid status by checking if `Order Status` is NOT "cancelled". If cancelled, deposit date is stripped.
- **Lazada**: Checks if status is "canceled". If status is "delivered" or "confirmed", it's treated as paid and the `deliveredDate` becomes the deposit date.

---

## 5. Calculation Examples

### Shopify PayMongo Calculation
- **Source Total**: 2500.00
- **Gateway**: "Paymongo GCash"
- **Output Gross**: Subtotal + Discount
- **Output Less Paymongo Fee**: `2500.00 * 0.015` = 37.50

### Shopee Multi-Item Grouping
- **Row 1**: Total Buyer Payment: 500, Service Fee: 10, Grand Total: 1000
- **Row 2**: Total Buyer Payment: 500, Service Fee: 10, Grand Total: 1000
- **Output Gross**: `500 + 500` = 1000
- **Output Less payment fee**: `10 + 10` = 20
- **Output Net**: `1000` (taken from Row 1 Grand Total)

---

## 6. Troubleshooting Guide

- **Missing `orders` sheet for Shopee**: Ensure the uploaded Shopee file is an `.xlsx` file and the active sheet contains headers like `Order ID` and `Grand Total`.
- **Unknown SKUs appear in `Unmapped SKUs`**: A new product was launched. Fix this by adding the product exact name to the `Offtake Report Summary per sku` sheet in the template, OR ask the developer to update `core/sku_normalizer.py`.
- **Formulas do not update until opening Excel**: Openpyxl does not compute Excel formulas. When you open the generated file, Excel will recalculate `% contribution` and `vs PY` automatically.
- **Dates under wrong month**: Ensure your CSV/XLSX exports use standard date formats (`YYYY-MM-DD`). The parsers in `core/workbook_utils.py` rely on this structure.

---

## 7. Developer Notes

| Area | File | Function/Class | Purpose |
|---|---|---|---|
| Shopify Parsing | `core/shopify_parser.py` | `process_shopify` | Merges orders and transactions CSVs |
| Shopee Parsing | `core/shopee_parser.py` | `load_shopee` | Groups items by `Order ID` |
| Lazada Parsing | `core/lazada_parser.py` | `load_lazada` | Groups items by `orderNumber` |
| SKU Normalization | `core/sku_normalizer.py` | `normalize_sku` | Hardcoded substring mappings |
| Valid SKU Loading | `core/sku_normalizer.py` | `extract_valid_skus_from_template` | Scrapes template for allowed SKUs |
| Report Generation | `core/report_generator.py` | `generate_full_report` | Main orchestrator to build the output XLSX |
| Safe Float Parsing | `core/workbook_utils.py` | `fval` | Safely casts empty/hyphen cells to `0.0` |

All spreadsheet logic relies on the `openpyxl` library. Formatting is copied cell-by-cell using `copy_style` to perfectly preserve the aesthetic of the original template.
