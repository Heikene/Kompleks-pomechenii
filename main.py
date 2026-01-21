# main.py
import sys
from PySide6.QtWidgets import QApplication
from ui.main_window import MainWindow

def main():
    app = QApplication(sys.argv)
    wnd = MainWindow()
    wnd.resize(800, 700)
    wnd.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
