import os
from PySide6.QtWidgets import QFrame, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QCheckBox, QFileDialog

class OutputSettingsPanel(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("OutputSettings")
        self.setStyleSheet("""
            QFrame#OutputSettings {
                background-color: #FFFFFF;
                border: 1px solid #E5E7EB;
                border-radius: 8px;
            }
            QLineEdit {
                border: 1px solid #D1D5DB;
                border-radius: 4px;
                padding: 6px;
                font-size: 13px;
            }
            QPushButton {
                background-color: #FFFFFF;
                border: 1px solid #D1D5DB;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: bold;
                color: #374151;
            }
            QPushButton:hover {
                background-color: #F3F4F6;
            }
            QCheckBox {
                font-size: 13px;
                color: #374151;
                padding: 4px 0px;
            }
        """)
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)
        
        title = QLabel("Output Settings")
        title.setStyleSheet("font-weight: 600; font-size: 16px; color: #111827; padding-bottom: 5px;")
        title.setMinimumHeight(24)
        self.main_layout.addWidget(title)
        
        # Folder row
        folder_layout = QVBoxLayout()
        folder_layout.setContentsMargins(0, 0, 0, 0)
        folder_layout.setSpacing(4)
        lbl_folder = QLabel("Output Folder")
        lbl_folder.setStyleSheet("font-size: 12px; color: #6B7280; padding-bottom: 2px;")
        lbl_folder.setMinimumHeight(16)
        
        folder_input_row = QHBoxLayout()
        folder_input_row.setContentsMargins(0, 0, 0, 0)
        self.folder_input = QLineEdit()
        self.folder_input.setText(os.getcwd())
        btn_browse = QPushButton("Browse")
        btn_browse.clicked.connect(self.browse_folder)
        
        folder_input_row.addWidget(self.folder_input)
        folder_input_row.addWidget(btn_browse)
        
        folder_layout.addWidget(lbl_folder)
        folder_layout.addLayout(folder_input_row)
        self.main_layout.addLayout(folder_layout)
        
        # Filename row
        fname_layout = QVBoxLayout()
        fname_layout.setContentsMargins(0, 0, 0, 0)
        fname_layout.setSpacing(4)
        lbl_fname = QLabel("Output Filename")
        lbl_fname.setStyleSheet("font-size: 12px; color: #6B7280;")
        
        self.fname_input = QLineEdit()
        self.fname_input.setText("Full_Sales_Report.xlsx")
        
        fname_layout.addWidget(lbl_fname)
        fname_layout.addWidget(self.fname_input)
        self.main_layout.addLayout(fname_layout)
        
        # Checkboxes
        self.chk_timestamp = QCheckBox("Add timestamp if file exists")
        self.chk_timestamp.setChecked(True)
        
        self.chk_open_folder = QCheckBox("Open folder after generation")
        self.chk_open_folder.setChecked(True)
        
        self.main_layout.addWidget(self.chk_timestamp)
        self.main_layout.addWidget(self.chk_open_folder)
        
        self.main_layout.addStretch()

    def browse_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if path:
            self.folder_input.setText(path)

    def get_output_path(self):
        folder = self.folder_input.text().strip()
        fname = self.fname_input.text().strip()
        if not fname.endswith(".xlsx"):
            fname += ".xlsx"
        return os.path.join(folder, fname)
