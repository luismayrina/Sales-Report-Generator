from PySide6.QtWidgets import QWidget, QFrame, QVBoxLayout, QHBoxLayout, QLabel

class ValidationChecklist(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("ValidationChecklist")
        self.setStyleSheet("""
            QFrame#ValidationChecklist {
                background-color: #FFFFFF;
                border: 1px solid #E5E7EB;
                border-radius: 8px;
            }
        """)
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)
        
        title = QLabel("Required Files")
        title.setStyleSheet("font-weight: 600; font-size: 16px; color: #111827; padding-bottom: 4px;")
        self.main_layout.addWidget(title)
        
        # Store items by key for easy updating
        self.items = {}
        
        self.add_item("shopify_orders", "Shopify Orders CSV")
        self.add_item("shopify_txns", "Shopify Transactions CSV")
        self.add_item("shopee", "Shopee Transaction Report")
        self.add_item("lazada", "Lazada Transaction Report")
        self.add_item("template", "Sample Reports Template")
        
        self.main_layout.addStretch()

    def add_item(self, key, name):
        row = QHBoxLayout()
        row.setSpacing(10)
        
        icon_lbl = QLabel("⚠️")
        icon_lbl.setStyleSheet("font-size: 14px;")
        icon_lbl.setFixedWidth(20)
        
        text_vbox = QVBoxLayout()
        text_vbox.setContentsMargins(0, 0, 0, 0)
        text_vbox.setSpacing(4)
        
        name_lbl = QLabel(name)
        name_lbl.setStyleSheet("font-size: 13px; font-weight: 500; color: #374151;")
        name_lbl.setWordWrap(True)
        
        status_lbl = QLabel("Missing")
        status_lbl.setStyleSheet("font-size: 11px; color: #F59E0B;") # Amber warning
        status_lbl.setWordWrap(True)
        
        text_vbox.addWidget(name_lbl)
        text_vbox.addWidget(status_lbl)
        
        row.addWidget(icon_lbl)
        row.addLayout(text_vbox)
        row.addStretch()
        
        self.main_layout.addLayout(row)
        
        self.items[key] = {
            "icon": icon_lbl,
            "status": status_lbl,
            "name_lbl": name_lbl
        }

    def update_item(self, key, path, status):
        if key not in self.items: return
        item = self.items[key]
        
        if status == "Valid":
            item["icon"].setText("✅")
            item["status"].setText("Ready")
            item["status"].setStyleSheet("font-size: 11px; color: #16A34A;") # Green
        elif status == "Invalid":
            item["icon"].setText("❌")
            item["status"].setText("Invalid file type")
            item["status"].setStyleSheet("font-size: 11px; color: #DC2626;") # Red
        else: # Empty
            item["icon"].setText("⚠️")
            item["status"].setText("Missing")
            item["status"].setStyleSheet("font-size: 11px; color: #F59E0B;") # Amber
            
    def all_valid(self):
        return all(item["status"].text() == "Ready" for item in self.items.values())
