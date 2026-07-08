# Sales Report Generator - User Guide

Welcome to the Sales Report Generator application. This desktop application consolidates sales data from e-commerce platforms (Shopify, Shopee, Lazada) and physical store channels into a unified, formatted Excel report.

This guide outlines the required files, their naming conventions, and how to successfully use the application.

---

## 1. How to Run the Application

To run the application, you can download the pre-built standalone version for your operating system (no setup or Python installation required).

### Downloading the App
1. Go to the **Releases** section of the GitHub repository.
2. Download the appropriate file for your system:
   * **For Windows:** Download **`Sales Report Generator.exe`**.
   * **For macOS:** Download **`Sales Report Generator.app.zip`** (extract the ZIP file to get the app).

### Launching the App
* **On Windows:** Double-click the downloaded **`Sales Report Generator.exe`** file.
* **On macOS:** Double-click the extracted **`Sales Report Generator.app`** bundle. 
  *(Note: The first time you open it, you may need to right-click the app and select **Open** to bypass macOS gatekeeper warnings for unsigned applications).*

---

## 2. Using the Application Interface

1. **Launch the App:** Run the Python script or open the built executable.
2. **Upload Files:** In the application interface, drag and drop the required files into their respective slots, or click the slots to browse your computer.
3. **Check Validation:** On the right side, the **Validation Checklist** will show green checkmarks as you upload the correct files.
4. **Set Output Preferences:** Choose where you want the final generated report to be saved.
5. **Generate:** Once all required files are uploaded and validated, click the **🚀 Generate Report** button. 
6. **View Results:** The application will process your data and open the folder containing your brand new `Full_Sales_Report.xlsx`.

---

## 3. Required Files and Formats

To successfully generate the report, you must provide the following files to the application:

### A. E-Commerce Platforms

| Slot Name | Expected Format | Naming Convention / Rules | Description |
| :--- | :--- | :--- | :--- |
| **Shopify Orders** | `.csv` | Usually `orders_export.csv` | Exported directly from your Shopify admin panel. Contains order-level data. |
| **Shopify Transactions** | `.csv` | Usually `transactions_export.csv` | Exported from Shopify. Used to calculate exact transaction fees (like Paymongo). |
| **Shopee Report** | `.xlsx` (Excel) | Usually `Shopee Transaction Report.xlsx` | The official transaction export from Shopee Seller Center. |
| **Lazada Report** | `.xlsx` (Excel) | Usually `Lazada Transaction Report.xlsx` | The official transaction export from Lazada Seller Center. |

### B. Report Template

| Slot Name | Expected Format | Naming Convention / Rules | Description |
| :--- | :--- | :--- | :--- |
| **Sample Reports Template** | `.xlsx` (Excel) | e.g., `Sample Reports.xlsx` | This is a blank or existing template file. The app reads this file to understand your formatting, styling, formulas, and SKU lists, and copies them to the new report. |

### C. Physical Stores & TikTok (Offtake Report)

Physical stores are handled slightly differently. You need the main Offtake Report template, and an **optional (but highly recommended)** folder containing the individual raw store reports.

#### 1. Offtake Report Excel (Required)
* **Format:** `.xlsx`
* **Naming:** E.g., `2026 OFFTAKE REPORT.xlsx`. The application looks for the year (e.g., 2026) in the file name to determine the current reporting year.
* **Purpose:** The app reads this file to pull historical **Previous Year (PY) data** and load any manual store entries (like Pick-a-roo) that aren't handled by external files.

#### 2. External Reports Folder (Optional)
* **Format:** A **Folder** on your computer.
* **Purpose:** Put all your raw store reports (PDFs, Excel, HTML) inside this folder. Upload the *entire folder* into the application. The app will scan the folder and automatically parse the stores based on keywords in the file names.

**Supported External Reports & Naming Rules:**
For the app to detect these files inside the folder, the file names **MUST** contain specific keywords (case-insensitive):

| Store Name | Required Format | Required Keyword in Filename | Needs Month in Filename? | Example Filename |
| :--- | :--- | :--- | :--- | :--- |
| **TikTok** | `.xlsx` | `tiktok` | ❌ No (Auto-detected from data) | `Tiktok Transaction Report.xlsx` |
| **Simula PH** | `.xlsx` | `simula` | ❌ No (Auto-detected from data) | `2026 SIMULA PH SALES REPORT.xlsx` |
| **Frankie and Friends** | `.xlsx` | `frankie` | ❌ No (Auto-detected from data) | `Frankie and Friends Sales Report.xlsx` |
| **Craft Central** | `.html` | `Craft Central` (Case sensitive) | ✅ Optional (Reads filename first, falls back to HTML text) | `The Craft Central Sales Report.html` |
| **9 Matters** | `.pdf` | `9matters` | ✅ **Yes** (Relies on filename, e.g., "April" or "Apr") | `9MATTERS-April Sales Report.pdf` |
| **Common Room** | `.pdf` | `common room` | ✅ **Yes** (Relies on filename, e.g., "Jan" or "January") | `Common Room [Jan] Sales Report.pdf` |

> [!IMPORTANT]
> **Simula PH Note:** Make sure your Simula PH report is saved as an Excel `.xlsx` file, not a PDF. The application reads the exact order number from the **Remarks** column.
> 
> **Month Detection Details:** When reading months from filenames (like for PDFs), the application understands both full names (e.g. `January`) and 3-letter abbreviations (e.g. `Jan`). It is not case-sensitive.

---

## 4. Frequently Asked Questions

**Q: Do I need to clean the Shopify CSVs before uploading?**
No. The application is designed to ingest the raw, unedited exports straight from Shopify.

**Q: What if I don't have an external report for a specific physical store yet?**
If a file (like TikTok or Simula PH) is missing from your External Reports Folder, the app will automatically try to read its data from the `[Year] OFFTAKE REPORT.xlsx` file as a fallback. 

**Q: How does the application match SKUs in the summary?**
The app reads the "Offtake Report Summary per sku" sheet inside your **Report Template**. It looks at your standardized SKU names in column A. When processing sales, it intelligently normalizes and matches product variations to ensure items are grouped correctly.

**Q: Does this work on Windows and Mac?**
Yes. The application and all required files are completely cross-platform.
