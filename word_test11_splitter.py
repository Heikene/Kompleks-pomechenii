# word_test11_splitter.py
from __future__ import annotations

import re
from copy import deepcopy
from typing import Optional

from docx.document import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table


def _norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ").replace("\u202F", " ").replace("ё", "е")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def _tbl_looks_like_test11_results(tbl: Table) -> bool:
    """
    Очень мягкая проверка, чтобы не резать "левые" таблицы:
    ищем в таблице признаки шапки результатов теста 11.
    """
    if not tbl.rows:
        return False
    # проверим первые ~10 строк на ключевые слова
    probe_rows = tbl.rows[: min(10, len(tbl.rows))]
    txt = " ".join(_norm(c.text) for r in probe_rows for c in r.cells)
    return ("результаты испытания" in txt) and ("фильтр" in txt) and ("расход" in txt)


def _set_repeat_header_rows_on_tbl_el(tbl_el, header_rows: int) -> None:
    """
    Ставит флаг "повторять как заголовок" на первые header_rows строк XML-таблицы.
    Работает напрямую с XML, без необходимости получать объект Table.
    """
    trs = tbl_el.xpath("./w:tr")
    for tr in trs[: max(0, header_rows)]:
        trPr = tr.find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            tr.insert(0, trPr)

        if trPr.find(qn("w:tblHeader")) is None:
            trPr.append(OxmlElement("w:tblHeader"))


def split_after_results_and_repeat_header(
    doc: DocxDocument,
    *,
    header_rows: int = 2,
    results_row_text: str = "Результаты испытания",
) -> bool:
    """
    Находит таблицу результатов "Проверка расхода приточного воздуха",
    ищет строку "Результаты испытания" и РЕЖЕТ таблицу так, что
    строка СРАЗУ ПОСЛЕ неё становится ПЕРВОЙ строкой новой таблицы.

    Далее ставит повтор заголовка (tblHeader) на первые header_rows строк новой таблицы.

    Возвращает True если разрезали, иначе False.
    """

    target_tbl: Optional[Table] = None
    results_row_idx: Optional[int] = None

    for t in doc.tables:
        if not _tbl_looks_like_test11_results(t):
            continue

        # ищем строку "Результаты испытания"
        for ri, row in enumerate(t.rows):
            row_text = _norm(" ".join(c.text for c in row.cells))
            if _norm(results_row_text) == row_text or _norm(results_row_text) in row_text:
                target_tbl = t
                results_row_idx = ri
                break

        if target_tbl is not None:
            break

    if target_tbl is None or results_row_idx is None:
        return False

    split_row_idx = results_row_idx + 1  # <-- строка после "Результаты испытания"
    if split_row_idx >= len(target_tbl.rows):
        return False

    tbl_el = target_tbl._tbl
    tr_elems = tbl_el.xpath("./w:tr")
    if len(tr_elems) != len(target_tbl.rows):
        return False

    # создаём новую таблицу как копию
    new_tbl_el = deepcopy(tbl_el)
    new_tr_elems = new_tbl_el.xpath("./w:tr")

    # в новой таблице удаляем всё ДО split_row_idx
    for i in range(split_row_idx - 1, -1, -1):
        new_tbl_el.remove(new_tr_elems[i])

    # в старой таблице удаляем всё Начиная с split_row_idx
    for i in range(len(tr_elems) - 1, split_row_idx - 1, -1):
        tbl_el.remove(tr_elems[i])

    # вставляем новую таблицу сразу после старой
    tbl_el.addnext(new_tbl_el)

    # делаем повтор шапки на новой таблице (теперь шапка = первые строки)
    _set_repeat_header_rows_on_tbl_el(new_tbl_el, header_rows)

    return True
