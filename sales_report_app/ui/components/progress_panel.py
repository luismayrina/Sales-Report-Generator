from PySide6.QtWidgets import QFrame, QVBoxLayout, QLabel, QProgressBar, QPlainTextEdit

class ProgressPanel(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("ProgressPanel")
        self.setStyleSheet("""
            QFrame#ProgressPanel {
                background-color: #FFFFFF;
                border: 1px solid #E5E7EB;
                border-radius: 8px;
            }
            QProgressBar {
                border: 1px solid #E5E7EB;
                border-radius: 4px;
                background-color: #F3F4F6;
                text-align: center;
                color: #111827;
                font-weight: bold;
                height: 18px;
            }
            QProgressBar::chunk {
                background-color: #2563EB;
                border-radius: 3px;
            }
            QPlainTextEdit {
                background-color: #1E1E1E;
                color: #D4D4D4;
                font-size: 11px;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(10)
        
        title = QLabel("Generation Progress")
        title.setStyleSheet("font-weight: 600; font-size: 16px; color: #111827;")
        self.main_layout.addWidget(title)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.main_layout.addWidget(self.progress_bar)
        
        self.log_console = QPlainTextEdit()
        self.log_console.setReadOnly(True)
        self.log_console.setMinimumHeight(150)
        self.main_layout.addWidget(self.log_console)

    def log(self, msg):
        self.log_console.appendPlainText(msg)

    def set_progress(self, val):
        self.progress_bar.setValue(val)
        
    def reset(self):
        self.progress_bar.setValue(0)
        self.log_console.clear()
