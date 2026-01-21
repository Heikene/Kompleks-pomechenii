from PySide6.QtWidgets import QApplication

def apply_compact_style(app: QApplication) -> None:
    app.setStyle("Fusion")
    app.setStyleSheet("""
        QWidget { font-size: 10pt; }
        QLineEdit, QComboBox, QDateEdit {
            padding: 5px 8px;
            border-radius: 6px;
        }
        QPushButton {
            padding: 6px 10px;
            border-radius: 6px;
        }
        QToolButton {
            padding: 4px 6px;
            border-radius: 6px;
        }
        QGroupBox {
            margin-top: 10px;
            font-weight: 600;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 4px;
        }
    """)
