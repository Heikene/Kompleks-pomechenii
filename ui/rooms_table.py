# ui/rooms_table.py  (новый файл)
from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QApplication
from PySide6.QtCore import Qt

class SpreadsheetTable(QTableWidget):
    """
    QTableWidget с excel-копипастой:
      • Ctrl+C — копирует текущее выделение как TSV
      • Ctrl+V — вставляет прямоугольный блок (TSV/CSV/строки)
      • если нет выделения — работаем от (0,0)
      • автоматически добавляет недостающие строки
    """
    def keyPressEvent(self, e):
        ctrl = e.modifiers() & Qt.ControlModifier
        if ctrl and e.key() == Qt.Key_C:
            self._copy_selection_to_clipboard()
            return
        if ctrl and e.key() == Qt.Key_V:
            self._paste_from_clipboard()
            return
        super().keyPressEvent(e)

    # --- helpers ---
    def _copy_selection_to_clipboard(self):
        rngs = self.selectedRanges()
        if rngs:
            r = rngs[0]
            rows = range(r.topRow(), r.bottomRow() + 1)
            cols = range(r.leftColumn(), r.rightColumn() + 1)
        else:
            # нет выделения — копируем всё содержимое
            rows = range(0, self.rowCount())
            cols = range(0, self.columnCount())

        lines = []
        for i in rows:
            vals = []
            for j in cols:
                it = self.item(i, j)
                vals.append("" if it is None else it.text())
            lines.append("\t".join(vals))
        QApplication.clipboard().setText("\n".join(lines))

    def _paste_from_clipboard(self):
        text = QApplication.clipboard().text()
        if not text:
            return

        # поддержка TSV/CSV
        rows_data = [
            [c.strip() for c in row.replace(";", "\t").split("\t")]
            for row in text.splitlines()
            if row.strip() != ""
        ]
        if not rows_data:
            return

        # точка вставки — левый верх выделения, иначе (0,0)
        rngs = self.selectedRanges()
        start_row = rngs[0].topRow() if rngs else 0
        start_col = rngs[0].leftColumn() if rngs else 0

        needed_rows = start_row + len(rows_data)
        if self.rowCount() < needed_rows:
            self.setRowCount(needed_rows)

        needed_cols = start_col + max(len(r) for r in rows_data)
        if self.columnCount() < needed_cols:
            self.setColumnCount(needed_cols)

        for i, row_vals in enumerate(rows_data):
            for j, val in enumerate(row_vals):
                r = start_row + i
                c = start_col + j
                it = self.item(r, c)
                if it is None:
                    it = QTableWidgetItem()
                    self.setItem(r, c, it)
                it.setText(val)