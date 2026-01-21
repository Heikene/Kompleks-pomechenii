from PySide6.QtWidgets import QMainWindow, QLabel
from PySide6.QtCore import Qt

class OldCleanroomWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Старый режим — Чистые помещения")
        label = QLabel("Режим чистых помещений пока реализован в старом виде.")
        label.setAlignment(Qt.AlignCenter)
        self.setCentralWidget(label)


def launch_old_app():
    return OldCleanroomWindow()
