# word_repeat_headers.py
from __future__ import annotations

import re
from copy import deepcopy
from typing import Optional

from docx.document import Document as DocxDocument
from docx.table import Table
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def _norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = (
        s.replace("\u00A0", " ")
         .replace("\u202F", " ")
         .replace("\u200B", "")
         .replace("\u00AD", "")
    )
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s.replace("ё", "е")


def _row_text(row) -> str:
    return _norm(" ".join(c.text for c in row.cells))


def _mark_tr_as_header(tr_el) -> None:
    trPr = tr_el.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr_el.insert(0, trPr)

    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def split_test_results_table(
    doc: DocxDocument,
    *,
    split_phrase: str = "Результаты испытания",
    header_rows: int = 2,
    table_must_contain: Optional[str] = None,
) -> int:
    """
    Разрезает таблицы так, чтобы строка ПОСЛЕ 'split_phrase' стала первой строкой новой таблицы.
    Затем помечает первые header_rows строк новой таблицы как повторяемый заголовок (tblHeader).

    Возвращает количество разрезанных таблиц.
    """
    split_key = _norm(split_phrase)
    must_key = _norm(table_must_contain) if table_must_contain else None

    changed = 0

    for t in doc.tables:
        if not t.rows:
            continue

        if must_key:
            whole = _norm(" ".join(_row_text(r) for r in t.rows[: min(len(t.rows), 6)]))
            if must_key not in whole:
                continue

        marker_idx = None
        for i, row in enumerate(t.rows):
            if split_key in _row_text(row):
                marker_idx = i
                break

        if marker_idx is None:
            continue

        start_new = marker_idx + 1
        if start_new >= len(t.rows):
            continue  # нечего переносить

        # XML таблицы и строки
        tbl_el = t._tbl
        tr_elems = tbl_el.xpath("./w:tr")
        if len(tr_elems) != len(t.rows):
            continue

        # Копия таблицы -> будущая "новая таблица"
        new_tbl_el = deepcopy(tbl_el)
        new_tr_elems = new_tbl_el.xpath("./w:tr")

        # В НОВОЙ таблице оставляем строки начиная с start_new
        for i in range(len(new_tr_elems) - 1, -1, -1):
            if i < start_new:
                new_tbl_el.remove(new_tr_elems[i])

        # В СТАРОЙ таблице оставляем строки до start_new-1 (включая 'Результаты испытания')
        for i in range(len(tr_elems) - 1, -1, -1):
            if i >= start_new:
                tbl_el.remove(tr_elems[i])

        # Пометить первые header_rows строк новой таблицы как повторяемый заголовок
        new_tr_elems2 = new_tbl_el.xpath("./w:tr")
        for i in range(min(header_rows, len(new_tr_elems2))):
            _mark_tr_as_header(new_tr_elems2[i])

        # Вставить новую таблицу сразу после старой
        tbl_el.addnext(new_tbl_el)

        changed += 1

    return changed
