import os
import sys
import subprocess
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QScrollArea, QPushButton, QMessageBox, QLabel
)
from PySide6.QtCore import Qt

# Import core logic
from core.report_generator import generate_full_report
from core.shopify_parser import process_shopify
from core.shopee_parser import load_shopee
from core.lazada_parser import load_lazada
from PySide6.QtCore import QThread, Signal

# Import UI components
from ui.components.upload_slot import UploadSlot
from ui.components.platform_card import PlatformCard
from ui.components.validation_checklist import ValidationChecklist
from ui.components.output_settings import OutputSettingsPanel
from ui.components.progress_panel import ProgressPanel

class WorkerThread(QThread):
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    def __init__(self, paths, output_path):
        super().__init__()
        self.paths = paths
        self.output_path = output_path

    def run(self):
        try:
            self.log_signal.emit("📂 Validating files...")
            self.progress_signal.emit(5)
            
            self.log_signal.emit("📂 Loading Shopify transactions & orders...")
            self.progress_signal.emit(10)
            shopify_orders = process_shopify(self.paths['shopify_orders'], self.paths['shopify_txns'])
            
            self.log_signal.emit("📂 Loading Shopee transactions...")
            self.progress_signal.emit(30)
            shopee_orders = load_shopee(self.paths['shopee'])
            
            self.log_signal.emit("📂 Loading Lazada transactions...")
            self.progress_signal.emit(50)
            lazada_orders = load_lazada(self.paths['lazada'])
            
            all_orders = shopify_orders + shopee_orders + lazada_orders
            self.log_signal.emit(f"📊 Total combined orders: {len(all_orders)}")
            self.progress_signal.emit(60)
            
            def thread_logger(msg):
                self.log_signal.emit(msg)

            success, result_path = generate_full_report(
                all_orders=all_orders,
                template_path=self.paths['template'],
                output_path=self.output_path,
                logger=thread_logger
            )
            
            self.progress_signal.emit(100)
            self.finished_signal.emit(success, result_path)

        except Exception as e:
            self.log_signal.emit(f"❌ Error during generation: {str(e)}")
            self.finished_signal.emit(False, str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sales Report Generator (Redesign V2)")
        self.setMinimumSize(1100, 720)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F6F7F9;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollArea > QWidget > QWidget {
                background-color: transparent;
            }
        """)

        # Main Central Widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Header
        header = QWidget()
        header.setStyleSheet("background-color: #FFFFFF; border-bottom: 1px solid #E5E7EB;")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(30, 20, 30, 20)
        
        title = QLabel("Sales Report Generator")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #111827;")
        subtitle = QLabel("Generate consolidated Shopify, Shopee, and Lazada sales reports from uploaded CSV and Excel files.")
        subtitle.setStyleSheet("font-size: 14px; color: #6B7280;")
        
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        main_layout.addWidget(header)

        # Content Area (Two Columns)
        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(30, 30, 30, 30)
        content_layout.setSpacing(30)
        main_layout.addLayout(content_layout)

        # ── LEFT COLUMN (Scrollable Platform Cards) ──
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(20)
        left_layout.setContentsMargins(0, 0, 10, 0)

        # State Dictionary to hold paths
        self.file_paths = {
            "shopify_orders": "",
            "shopify_txns": "",
            "shopee": "",
            "lazada": "",
            "template": ""
        }

        # Shopify Card
        shopify_card = PlatformCard("Shopify Reports", "Upload your Shopify order and transaction exports.")
        s1 = UploadSlot("shopify_orders", "Shopify Orders CSV", "Usually named orders_export.csv", "CSV Files (*.csv)")
        s2 = UploadSlot("shopify_txns", "Shopify Transactions CSV", "Usually named transactions_export.csv", "CSV Files (*.csv)")
        shopify_card.add_slot(s1)
        shopify_card.add_slot(s2)
        left_layout.addWidget(shopify_card)

        # Shopee Card
        shopee_card = PlatformCard("Shopee Report", "Upload the Shopee transaction report. The app reads the orders sheet.")
        s3 = UploadSlot("shopee", "Shopee Transaction Report", "Accepted format: .xlsx", "Excel Files (*.xlsx)")
        shopee_card.add_slot(s3)
        left_layout.addWidget(shopee_card)

        # Lazada Card
        lazada_card = PlatformCard("Lazada Report", "Upload the Lazada transaction report.")
        s4 = UploadSlot("lazada", "Lazada Transaction Report", "Accepted format: .xlsx", "Excel Files (*.xlsx)")
        lazada_card.add_slot(s4)
        left_layout.addWidget(lazada_card)

        # Template Card
        template_card = PlatformCard("Report Template", "Upload the Sample Reports.xlsx template used for formatting, formulas, and report structure.")
        s5 = UploadSlot("template", "Sample Reports Template", "Accepted format: .xlsx", "Excel Files (*.xlsx)")
        template_card.add_slot(s5)
        left_layout.addWidget(template_card)

        left_layout.addStretch()
        left_scroll.setWidget(left_widget)
        content_layout.addWidget(left_scroll, stretch=5)

        # ── RIGHT COLUMN (Scrollable Panels) ──
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_inner = QWidget()
        right_layout = QVBoxLayout(right_inner)
        right_layout.setSpacing(20)
        right_layout.setContentsMargins(0, 0, 10, 0)

        self.checklist = ValidationChecklist()
        right_layout.addWidget(self.checklist)

        self.settings_panel = OutputSettingsPanel()
        right_layout.addWidget(self.settings_panel)

        self.progress_panel = ProgressPanel()
        right_layout.addWidget(self.progress_panel)

        self.btn_generate = QPushButton("Upload required files first")
        self.btn_generate.setFixedHeight(50)
        self.btn_generate.setStyleSheet("""
            QPushButton {
                background-color: #D1D5DB;
                color: #6B7280;
                font-weight: bold;
                font-size: 16px;
                border-radius: 8px;
                border: none;
            }
        """)
        self.btn_generate.setEnabled(False)
        self.btn_generate.clicked.connect(self.start_generation)
        right_layout.addWidget(self.btn_generate)
        
        right_layout.addStretch()
        right_scroll.setWidget(right_inner)

        content_layout.addWidget(right_scroll, stretch=5)

        # Connect Signals
        for slot in [s1, s2, s3, s4, s5]:
            slot.file_changed.connect(self.on_file_changed)

        self.last_output_path = None

    def on_file_changed(self, key, path, status):
        self.file_paths[key] = path
        self.checklist.update_item(key, path, status)
        
        # Check if all valid
        if self.checklist.all_valid():
            self.btn_generate.setEnabled(True)
            self.btn_generate.setText("🚀 Generate Report")
            self.btn_generate.setStyleSheet("""
                QPushButton {
                    background-color: #2563EB;
                    color: white;
                    font-weight: bold;
                    font-size: 16px;
                    border-radius: 8px;
                    border: none;
                }
                QPushButton:hover { background-color: #1D4ED8; }
                QPushButton:pressed { background-color: #1E40AF; }
            """)
        else:
            self.btn_generate.setEnabled(False)
            self.btn_generate.setText("Upload required files first")
            self.btn_generate.setStyleSheet("""
                QPushButton {
                    background-color: #D1D5DB;
                    color: #6B7280;
                    font-weight: bold;
                    font-size: 16px;
                    border-radius: 8px;
                    border: none;
                }
            """)

    def start_generation(self):
        self.btn_generate.setEnabled(False)
        self.btn_generate.setText("⏳ Generating...")
        self.progress_panel.reset()
        
        out_path = self.settings_panel.get_output_path()
        
        self.worker = WorkerThread(self.file_paths, out_path)
        self.worker.log_signal.connect(self.progress_panel.log)
        self.worker.progress_signal.connect(self.progress_panel.set_progress)
        self.worker.finished_signal.connect(self.on_generation_finished)
        self.worker.start()

    def on_generation_finished(self, success, path_or_err):
        if success:
            self.last_output_path = path_or_err
            self.btn_generate.setText("✅ Report Generated")
            self.btn_generate.setStyleSheet("""
                QPushButton {
                    background-color: #16A34A;
                    color: white;
                    font-weight: bold;
                    font-size: 16px;
                    border-radius: 8px;
                    border: none;
                }
                QPushButton:hover { background-color: #15803D; }
            """)
            
            QMessageBox.information(self, "Success", f"Report generated successfully!\n\nSaved to:\n{path_or_err}")
            
            if self.settings_panel.chk_open_folder.isChecked():
                folder = os.path.dirname(path_or_err)
                if sys.platform == "win32":
                    os.startfile(folder)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", folder])
                else:
                    subprocess.Popen(["xdg-open", folder])
                    
            # Reset button to allow another generation
            self.btn_generate.setEnabled(True)
            self.btn_generate.setText("🚀 Generate Another Report")
            self.btn_generate.setStyleSheet("""
                QPushButton {
                    background-color: #2563EB;
                    color: white;
                    font-weight: bold;
                    font-size: 16px;
                    border-radius: 8px;
                    border: none;
                }
                QPushButton:hover { background-color: #1D4ED8; }
            """)
        else:
            self.btn_generate.setEnabled(True)
            self.btn_generate.setText("🚀 Try Again")
            self.btn_generate.setStyleSheet("""
                QPushButton {
                    background-color: #2563EB;
                    color: white;
                    font-weight: bold;
                    font-size: 16px;
                    border-radius: 8px;
                    border: none;
                }
            """)
            QMessageBox.critical(self, "Generation Failed", f"An error occurred during generation:\n\n{path_or_err}\n\nPlease check the logs in the progress panel.")
