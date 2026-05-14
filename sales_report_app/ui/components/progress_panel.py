from PySide6.QtWidgets import QFrame, QVBoxLayout, QLabel, QProgressBar, QPlainTextEdit

class ProgressPanel(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("progressPanel")
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(10)
        
        title = QLabel("Generation Progress")
        title.setObjectName("cardTitle")
        self.main_layout.addWidget(title)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.main_layout.addWidget(self.progress_bar)
        
        self.log_console = QPlainTextEdit()
        self.log_console.setObjectName("progressLog")
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
