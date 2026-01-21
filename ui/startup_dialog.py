# ui/startup_dialog.py
from PySide6.QtWidgets import QDialog, QVBoxLayout, QPushButton

class StartupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выберите режим работы")
        self.resize(300, 150)
        self.choice = None

        layout = QVBoxLayout(self)

        btn_eq = QPushButton("Оборудование")
        btn_cr = QPushButton("Чистые помещения")

        btn_eq.clicked.connect(lambda: self._select("equipment"))
        btn_cr.clicked.connect(lambda: self._select("cleanrooms"))

        layout.addWidget(btn_eq)
        layout.addWidget(btn_cr)

    def _select(self, choice):
        self.choice = choice
        self.accept()

    def get_choice(self):
        return self.choice
