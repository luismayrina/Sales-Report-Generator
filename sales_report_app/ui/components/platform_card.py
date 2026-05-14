from PySide6.QtWidgets import QFrame, QVBoxLayout, QLabel

class PlatformCard(QFrame):
    def __init__(self, title, description, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        
        # Inner layout to respect padding
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(12)
        
        # Title
        self.title_label = QLabel(title)
        self.title_label.setObjectName("cardTitle")
        self.main_layout.addWidget(self.title_label)
        
        # Description
        if description:
            self.desc_label = QLabel(description)
            self.desc_label.setObjectName("cardSubtitle")
            self.desc_label.setWordWrap(True)
            self.main_layout.addWidget(self.desc_label)
            
        # VBox to hold the actual slots
        self.slots_layout = QVBoxLayout()
        self.slots_layout.setSpacing(12)
        self.main_layout.addLayout(self.slots_layout)

    def add_slot(self, slot_widget):
        self.slots_layout.addWidget(slot_widget)
