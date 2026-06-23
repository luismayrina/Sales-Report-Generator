from PySide6.QtWidgets import QWidget, QFrame, QVBoxLayout, QHBoxLayout, QLabel

class ValidationChecklist(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)
        
        title = QLabel("Required Files")
        title.setObjectName("cardTitle")
        self.main_layout.addWidget(title)
        
        # Store items by key for easy updating
        self.items = {}
        
        self.add_item("shopify_orders", "Shopify Orders CSV")
        self.add_item("shopify_txns", "Shopify Transactions CSV")
        self.add_item("shopee", "Shopee Transaction Report")
        self.add_item("lazada", "Lazada Transaction Report")
        self.add_item("offtake_report", "Offtake Report Excel", optional=True)
        self.add_item("template", "Sample Reports Template")
        self.add_item("external_docs", "External Reports Folder", optional=True)
        
        self.main_layout.addStretch()

    def add_item(self, key, name, optional=False):
        row = QHBoxLayout()
        row.setSpacing(10)
        
        icon_lbl = QLabel("⚠️" if not optional else "📁")
        icon_lbl.setStyleSheet("font-size: 14px;")
        icon_lbl.setFixedWidth(20)
        
        text_vbox = QVBoxLayout()
        text_vbox.setContentsMargins(0, 0, 0, 0)
        text_vbox.setSpacing(4)
        
        name_lbl = QLabel(name + (" (Optional)" if optional else ""))
        # Inherits default text color from QWidget
        name_lbl.setWordWrap(True)
        
        status_lbl = QLabel("Missing" if not optional else "Not Provided")
        if not optional:
            status_lbl.setObjectName("statusWarning")
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
            "name_lbl": name_lbl,
            "optional": optional
        }

    def update_item(self, key, path, status):
        if key not in self.items: return
        item = self.items[key]
        
        if status == "Valid":
            item["icon"].setText("✅")
            item["status"].setText("Ready")
            item["status"].setObjectName("statusReady")
        elif status == "Invalid":
            item["icon"].setText("❌")
            item["status"].setText("Invalid file type")
            item["status"].setObjectName("statusMissing")
        else: # Empty
            if item["optional"]:
                item["icon"].setText("📁")
                item["status"].setText("Not Provided")
                item["status"].setObjectName("") # default
            else:
                item["icon"].setText("⚠️")
                item["status"].setText("Missing")
                item["status"].setObjectName("statusWarning")
            
        # Force style update
        item["status"].style().unpolish(item["status"])
        item["status"].style().polish(item["status"])
            
    def all_valid(self):
        return all(
            item["status"].text() == "Ready" or (item["optional"] and item["status"].text() == "Not Provided")
            for item in self.items.values()
        )
