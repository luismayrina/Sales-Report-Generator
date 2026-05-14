import os
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog
from PySide6.QtCore import Signal, Qt

class UploadSlot(QWidget):
    # Signals emitted when file path changes or status changes
    file_changed = Signal(str, str, str) # key, path, status ("Valid", "Invalid", "Empty")

    def __init__(self, key, label_text, helper_text, file_filter="*.*", parent=None):
        super().__init__(parent)
        self.key = key
        self.file_filter = file_filter
        self.current_path = ""
        self.status = "Empty" # Empty, Valid, Invalid
        
        self.setAcceptDrops(True)
        
        # Main layout
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(4)
        
        # Helper text above the slot
        self.helper_label = QLabel(helper_text)
        self.helper_label.setStyleSheet("color: #6B7280; font-size: 12px;")
        self.main_layout.addWidget(self.helper_label)
        
        # The interactive drop zone box
        self.drop_zone = QWidget()
        self.drop_zone.setObjectName("dropZone")
        self.drop_zone_layout = QHBoxLayout(self.drop_zone)
        self.drop_zone_layout.setContentsMargins(15, 15, 15, 15)
        
        # Icon / Status Label
        self.status_icon = QLabel("📁")
        self.status_icon.setStyleSheet("font-size: 18px;")
        
        # Text area
        self.text_layout = QVBoxLayout()
        self.title_label = QLabel("Drop file here or browse")
        self.title_label.setStyleSheet("font-weight: bold; color: #111827; font-size: 13px;")
        self.title_label.setWordWrap(True)
        
        self.path_label = QLabel(label_text)
        self.path_label.setStyleSheet("color: #6B7280; font-size: 11px;")
        self.path_label.setWordWrap(True)
        
        self.text_layout.addWidget(self.title_label)
        self.text_layout.addWidget(self.path_label)
        
        # Buttons
        self.btn_browse = QPushButton("Browse")
        self.btn_browse.setCursor(Qt.PointingHandCursor)
        self.btn_browse.clicked.connect(self.browse_file)
        
        self.btn_remove = QPushButton("Remove")
        self.btn_remove.setCursor(Qt.PointingHandCursor)
        self.btn_remove.clicked.connect(self.remove_file)
        self.btn_remove.hide() # Hidden by default
        
        self.drop_zone_layout.addWidget(self.status_icon)
        self.drop_zone_layout.addLayout(self.text_layout, 1)
        self.drop_zone_layout.addWidget(self.btn_browse)
        self.drop_zone_layout.addWidget(self.btn_remove)
        
        self.main_layout.addWidget(self.drop_zone)
        self.update_style()

    def set_file(self, path):
        if not path or not os.path.exists(path):
            self.remove_file()
            return
            
        self.current_path = path
        filename = os.path.basename(path)
        
        # Basic validation (just extension for now, can be expanded)
        ext = os.path.splitext(filename)[1].lower()
        if ("csv" in self.file_filter.lower() and ext != ".csv") or \
           ("xlsx" in self.file_filter.lower() and ext != ".xlsx"):
            self.status = "Invalid"
            self.status_icon.setText("❌")
            self.title_label.setText(f"Invalid file: {filename}")
            self.title_label.setStyleSheet("font-weight: bold; color: #DC2626; font-size: 13px;")
        else:
            self.status = "Valid"
            self.status_icon.setText("✅")
            self.title_label.setText(filename)
            self.title_label.setStyleSheet("font-weight: bold; color: #16A34A; font-size: 13px;")
            
        self.path_label.setText(path)
        self.btn_browse.setText("Change")
        self.btn_remove.show()
        
        self.update_style()
        self.file_changed.emit(self.key, self.current_path, self.status)

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select File", "", self.file_filter)
        if path:
            self.set_file(path)

    def remove_file(self):
        self.current_path = ""
        self.status = "Empty"
        self.status_icon.setText("📁")
        self.title_label.setText("Drop file here or browse")
        self.title_label.setStyleSheet("font-weight: bold; color: #111827; font-size: 13px;")
        self.path_label.setText(self.helper_label.text()) # Reset to helper text
        self.btn_browse.setText("Browse")
        self.btn_remove.hide()
        
        self.update_style()
        self.file_changed.emit(self.key, self.current_path, self.status)

    def update_style(self):
        base_style = """
            QWidget#dropZone {
                background-color: #FAFAFA;
                border-radius: 6px;
            }
            QPushButton {
                background-color: #FFFFFF;
                border: 1px solid #D1D5DB;
                border-radius: 4px;
                padding: 4px 10px;
                color: #374151;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #F3F4F6;
            }
        """
        
        if self.status == "Empty":
            border = "border: 1px dashed #9CA3AF;"
        elif self.status == "Valid":
            border = "border: 1px solid #16A34A; background-color: #F0FDF4;"
        elif self.status == "Invalid":
            border = "border: 1px solid #DC2626; background-color: #FEF2F2;"
            
        self.drop_zone.setStyleSheet(base_style.replace("border-radius: 6px;", f"border-radius: 6px;\n{border}"))

    # Drag and Drop events
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_zone.setStyleSheet(self.drop_zone.styleSheet() + "QWidget#dropZone { border: 2px dashed #2563EB; background-color: #EFF6FF; }")

    def dragLeaveEvent(self, event):
        self.update_style()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            self.set_file(path)
