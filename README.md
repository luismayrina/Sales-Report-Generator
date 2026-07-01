# Sales Report Generator

A cross-platform desktop application designed to consolidate sales data from multiple e-commerce platforms and physical/offline stores into a single, unified, and formatted Excel report.

---

## 🚀 Quick Start for Clients

If you just want to run the application, you do **not** need to install Python or set up any code.

1. Go to the **[Releases](../../releases)** section of this GitHub repository.
2. Download the pre-built application for your operating system:
   * **Windows:** Download `Sales-Report-Generator-Windows.exe`
   * **macOS:** Download `Sales-Report-Generator-Mac.zip` (extract the ZIP to get the app bundle).
3. Open the application and follow the instructions in the interface.

For a detailed guide on required files, file formats, naming rules, and troubleshooting, please refer to the **[User Manual (user_manual.md)](sales_report_app/user_manual.md)**.
*(Note: If you are looking at the user manual inside the offline zip, it is located at `sales_report_app/user_manual.md`)*.

---

## 🛠️ For Developers

If you want to run the application from source code or modify the logic:

### Prerequisites
* Python 3.9 or higher installed.

### Installation & Running
1. Clone the repository.
2. Install the dependencies:
   ```bash
   pip install -r sales_report_app/requirements.txt
   ```
3. Run the GUI application:
   ```bash
   python sales_report_app/app.py
   ```

### Project Structure
* `sales_report_app/app.py` - Main application entry point (PySide6 GUI).
* `sales_report_app/core/` - Data parsers (Shopify, Shopee, Lazada, TikTok, physical stores) and report generation logic.
* `sales_report_app/ui/` - Layout and styling components for the user interface.
* `sales_report_app/build_scripts/` - Automated PyInstaller build scripts for Windows and macOS.

### Building Executables Locally
If you want to compile standalone executables locally on your machine:
* **Windows:** Run `sales_report_app/build_scripts/build_windows.bat`
* **macOS:** Run `sales_report_app/build_scripts/build_mac.sh`

---

## 🤖 Automated Builds (CI/CD)

This repository is equipped with an automated GitHub Actions pipeline (`.github/workflows/build.yml`). 

Whenever you push a version tag (e.g. `v1.0.0`) or publish a new release on GitHub, the system will automatically:
1. Spin up virtual environments for Windows and macOS.
2. Compile the source code into standalone executables.
3. Automatically attach `Sales-Report-Generator-Windows.exe` and `Sales-Report-Generator-Mac.zip` to the release.
