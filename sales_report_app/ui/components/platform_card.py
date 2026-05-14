from PySide6.QtWidgets import QFrame, QVBoxLayout, QLabel

class PlatformCard(QFrame):
    def __init__(self, title, description, parent=None):
        super().__init__(parent)
        self.setObjectName("PlatformCard")
        
        self.setStyleSheet("""
            QFrame#PlatformCard {
                background-color: #FFFFFF;
                border: 1px solid #E5E7EB;
                border-radius: 8px;
            }
        """)
        
        # Inner layout to respect padding
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(12)
        
        # Title
        self.title_label = QLabel(title)
        self.title_label.setStyleSheet("font-weight: 600; font-size: 16px; color: #111827;")
        self.main_layout.addWidget(self.title_label)
        
        # Description
        if description:
            self.desc_label = QLabel(description)
            self.desc_label.setWordWrap(True)
            self.desc_label.setStyleSheet("color: #6B7280; font-size: 13px; margin-bottom: 8px;")
            self.main_layout.addWidget(self.desc_label)
            
        # VBox to hold the actual slots
        self.slots_layout = QVBoxLayout()
        self.slots_layout.setSpacing(12)
        self.main_layout.addLayout(self.slots_layout)

    def add_slot(self, slot_widget):
        self.slots_layout.addWidget(slot_widget)
