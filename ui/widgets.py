# ui/widgets.py
from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QWidget, QHBoxLayout, QLineEdit, QToolButton, QFileDialog


class PathEdit(QWidget):
    textChanged = Signal(str)

    def __init__(self, parent=None, *, file_filter: str = "All files (*.*)", kind: str = "open"):
        super().__init__(parent)
        self._filter = file_filter
        self._kind = kind  # "open" | "save" | "dir"

        self._edit = QLineEdit(self)
        self._btn = QToolButton(self)
        self._btn.setText("…")
        self._btn.setCursor(Qt.PointingHandCursor)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        lay.addWidget(self._edit, 1)
        lay.addWidget(self._btn)

        self._edit.textChanged.connect(self.textChanged)
        self._btn.clicked.connect(self._browse)

    # --- QLineEdit-like API (важно для совместимости) ---
    def text(self) -> str:
        return self._edit.text()

    def setText(self, s: str) -> None:
        self._edit.setText(s)

    def setReadOnly(self, ro: bool) -> None:
        self._edit.setReadOnly(ro)

    def clear(self) -> None:
        self._edit.clear()

    def _browse(self) -> None:
        start_dir = os.path.dirname(self.text()) if self.text() else os.getcwd()

        if self._kind == "dir":
            p = QFileDialog.getExistingDirectory(self, "Выберите папку", start_dir)
            if p:
                self.setText(str(Path(p).resolve()))
            return

        if self._kind == "save":
            p, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", start_dir, self._filter)
        else:
            p, _ = QFileDialog.getOpenFileName(self, "Выберите файл", start_dir, self._filter)

        if p:
            self.setText(str(Path(p).resolve()))
