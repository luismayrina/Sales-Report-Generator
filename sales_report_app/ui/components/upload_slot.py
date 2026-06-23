import os
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog
from PySide6.QtCore import Signal, Qt

class UploadSlot(QWidget):
    # Signals emitted when file path changes or status changes
    file_changed = Signal(str, str, str) # key, path, status ("Valid", "Invalid", "Empty")

    def __init__(self, key, label_text, helper_text, file_filter="*.*", is_folder=False, parent=None):
        super().__init__(parent)
        self.key = key
        self.file_filter = file_filter
        self.is_folder = is_folder
        self.current_path = ""
        self.status = "Empty" # Empty, Valid, Invalid
        
        self.setAcceptDrops(True)
        
        # Main layout
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(4)
        
        # Helper text above the slot
        self.helper_label = QLabel(helper_text)
        self.helper_label.setObjectName("uploadSlotHint")
        self.main_layout.addWidget(self.helper_label)
        
        # The interactive drop zone box
        self.drop_zone = QWidget()
        self.drop_zone.setObjectName("uploadSlot")
        self.drop_zone_layout = QHBoxLayout(self.drop_zone)
        self.drop_zone_layout.setContentsMargins(15, 15, 15, 15)
        
        # Icon / Status Label
        self.status_icon = QLabel("📁")
        self.status_icon.setStyleSheet("font-size: 18px;")
        
        # Text area
        self.text_layout = QVBoxLayout()
        self.title_label = QLabel("Drop folder here or browse" if is_folder else "Drop file here or browse")
        self.title_label.setObjectName("uploadSlotLabel")
        self.title_label.setWordWrap(True)
        
        self.path_label = QLabel(label_text)
        self.path_label.setObjectName("uploadSlotHint")
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
        
        if self.is_folder:
            if not os.path.isdir(path):
                self.status = "Invalid"
                self.status_icon.setText("❌")
                self.title_label.setText(f"Not a folder: {filename}")
                self.title_label.setObjectName("statusMissing")
                self.title_label.style().unpolish(self.title_label)
                self.title_label.style().polish(self.title_label)
            else:
                self.status = "Valid"
                self.status_icon.setText("✅")
                self.title_label.setText(filename)
                self.title_label.setObjectName("statusReady")
                self.title_label.style().unpolish(self.title_label)
                self.title_label.style().polish(self.title_label)
        else:
            # Basic validation (just extension for now, can be expanded)
            ext = os.path.splitext(filename)[1].lower()
            if ("csv" in self.file_filter.lower() and ext != ".csv") or \
               ("xlsx" in self.file_filter.lower() and ext != ".xlsx"):
                self.status = "Invalid"
                self.status_icon.setText("❌")
                self.title_label.setText(f"Invalid file: {filename}")
                self.title_label.setObjectName("statusMissing")
                self.title_label.style().unpolish(self.title_label)
                self.title_label.style().polish(self.title_label)
            else:
                self.status = "Valid"
                self.status_icon.setText("✅")
                self.title_label.setText(filename)
                self.title_label.setObjectName("statusReady")
                self.title_label.style().unpolish(self.title_label)
                self.title_label.style().polish(self.title_label)
            
        self.path_label.setText(path)
        self.btn_browse.setText("Change")
        self.btn_remove.show()
        
        self.update_style()
        self.file_changed.emit(self.key, self.current_path, self.status)

    def browse_file(self):
        if self.is_folder:
            path = QFileDialog.getExistingDirectory(self, "Select Directory", "")
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Select File", "", self.file_filter)
        if path:
            self.set_file(path)

    def remove_file(self):
        self.current_path = ""
        self.status = "Empty"
        self.status_icon.setText("📁")
        self.title_label.setText("Drop folder here or browse" if self.is_folder else "Drop file here or browse")
        self.title_label.setObjectName("uploadSlotLabel")
        self.title_label.style().unpolish(self.title_label)
        self.title_label.style().polish(self.title_label)
        self.path_label.setText(self.helper_label.text()) # Reset to helper text
        self.btn_browse.setText("Browse")
        self.btn_remove.hide()
        
        self.update_style()
        self.file_changed.emit(self.key, self.current_path, self.status)

    def update_style(self):
        if self.status == "Empty":
            self.drop_zone.setStyleSheet("")
        elif self.status == "Valid":
            # Just force the border and background dynamically, overriding theme
            self.drop_zone.setStyleSheet("QWidget#uploadSlot { border: 1px solid #16A34A; background-color: rgba(22, 163, 74, 0.05); border-radius: 8px; }")
        elif self.status == "Invalid":
            self.drop_zone.setStyleSheet("QWidget#uploadSlot { border: 1px solid #DC2626; background-color: rgba(220, 38, 38, 0.05); border-radius: 8px; }")

    # Drag and Drop events
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_zone.setStyleSheet("QWidget#uploadSlot { border: 2px dashed #2563EB; background-color: rgba(59, 130, 246, 0.05); border-radius: 8px; }")

    def dragLeaveEvent(self, event):
        self.update_style()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            self.set_file(path)
