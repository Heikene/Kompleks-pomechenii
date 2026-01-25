import re
from copy import deepcopy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table

def _norm(s: str) -> str:
    s = (s or "").replace("\u00A0", " ").replace("\u202F", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def _set_repeat_header(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))

def split_after_results_and_repeat_header(doc, *, header_rows: int = 2) -> bool:
    """
    Делит таблицу так, чтобы:
      - строка "Результаты испытания" ОСТАЛАСЬ в первой таблице
      - следующая строка стала 1-й строкой новой таблицы
    И в новой таблице помечает первые header_rows строк как повторяемые заголовки.
    """
    target_tbl: Table | None = None
    results_row_idx: int | None = None

    # 1) найти таблицу и строку "Результаты испытания"
    for t in doc.tables:
        for ri, row in enumerate(t.rows):
            row_text = _norm(" ".join(c.text for c in row.cells))
            if row_text == _norm("Результаты испытания"):
                target_tbl = t
                results_row_idx = ri
                break
        if target_tbl:
            break

    if not target_tbl or results_row_idx is None:
        return False

    # 2) точка разреза = СЛЕДУЮЩАЯ строка после "Результаты испытания"
    cut_idx = results_row_idx + 1
    if cut_idx >= len(target_tbl.rows):
        return False  # нечего переносить

    tbl_el = target_tbl._tbl
    tr_elems = tbl_el.xpath("./w:tr")
    if len(tr_elems) != len(target_tbl.rows):
        return False

    # 3) новая таблица = копия всей, но оставляем строки [cut_idx:]
    new_tbl_el = deepcopy(tbl_el)
    new_trs = new_tbl_el.xpath("./w:tr")

    keep_new = set(range(cut_idx, len(new_trs)))
    for i in range(len(new_trs) - 1, -1, -1):
        if i not in keep_new:
            new_tbl_el.remove(new_trs[i])

    # 4) старая таблица = оставляем строки [:cut_idx] (то есть включая "Результаты испытания")
    for i in range(len(tr_elems) - 1, -1, -1):
        if i >= cut_idx:
            tbl_el.remove(tr_elems[i])

    # 5) вставить новую таблицу после старой
    target_tbl._tbl.addnext(new_tbl_el)

    # 6) пометить первые header_rows строк новой таблицы как повторяемые
    #    проще всего: снова найти таблицу, у которой первая строка == "Фильтр ..."
    for t in doc.tables:
        if not t.rows:
            continue
        first = _norm(" ".join(c.text for c in t.rows[0].cells))
        if "фильтр" in first and "скорость" in first and "расход" in first:
            for i in range(min(header_rows, len(t.rows))):
                _set_repeat_header(t.rows[i])
            break

    return True
