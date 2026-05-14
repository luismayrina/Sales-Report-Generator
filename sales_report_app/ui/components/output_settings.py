import os
from PySide6.QtWidgets import QFrame, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QCheckBox, QFileDialog

class OutputSettingsPanel(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)
        
        title = QLabel("Output Settings")
        title.setObjectName("cardTitle")
        title.setMinimumHeight(24)
        self.main_layout.addWidget(title)
        
        # Folder row
        folder_layout = QVBoxLayout()
        folder_layout.setContentsMargins(0, 0, 0, 0)
        folder_layout.setSpacing(4)
        lbl_folder = QLabel("Output Folder")
        lbl_folder.setObjectName("cardSubtitle")
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
        lbl_fname.setObjectName("cardSubtitle")
        
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
