# word_update_all.py
from __future__ import annotations

import re
from typing import Optional

import win32com.client as win32
from win32com.client import constants as c


def _clean_text(s: str) -> str:
    # Word часто возвращает текст с '\r\x07'
    s = s.replace("\r", "").replace("\x07", "")
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _norm(s: str) -> str:
    s = _clean_text(s).lower()
    return s.replace("ё", "е")


def _table_text_first_rows(tbl, rows: int = 3) -> str:
    """Собираем текст первых N строк таблицы одним куском."""
    rows = min(rows, int(tbl.Rows.Count))
    parts = []
    for i in range(1, rows + 1):
        try:
            r = tbl.Rows(i)
            for cell in r.Cells:
                parts.append(_clean_text(cell.Range.Text))
        except Exception:
            continue
    return " ".join(parts)


def _looks_like_test11_3_table(tbl) -> bool:
    """
    Узнаём таблицу "Тест 11.3 Проверка расхода приточного воздуха" по шапке:
    'Фильтр', 'Скорость потока', 'Расход приточного воздуха', 'Соответствует'
    (и обычно встречается 'ДА'/'НЕТ').
    """
    try:
        if int(tbl.Rows.Count) < 2:
            return False
        head = _norm(_table_text_first_rows(tbl, rows=3))
        return (
            ("фильтр" in head)
            and ("скорост" in head)          # скорость/скорости
            and ("расход" in head)
            and ("приточ" in head)           # приточного
            and ("соответ" in head)          # соответствует
        )
    except Exception:
        return False


def _set_repeat_header(tbl, header_rows_count: int) -> None:
    """
    Делает первые header_rows_count строк повторяемыми как шапку на каждой странице.
    """
    header_rows_count = max(1, header_rows_count)
    max_rows = int(tbl.Rows.Count)
    header_rows_count = min(header_rows_count, max_rows)

    # На всякий: чтобы Word не разрывал строки (особенно шапку) пополам
    try:
        tbl.Rows.AllowBreakAcrossPages = True  # данные можно, если нужно
    except Exception:
        pass

    for i in range(1, header_rows_count + 1):
        try:
            tbl.Rows(i).HeadingFormat = True
        except Exception:
            pass
        try:
            tbl.Rows(i).AllowBreakAcrossPages = False
        except Exception:
            pass


def enforce_test11_3_header_each_page(docx_path: str, header_rows_count: int = 2) -> int:
    """
    Находит все таблицы формата Тест 11.3 и включает повтор шапки на каждой странице.
    Возвращает количество найденных/исправленных таблиц.
    """
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False

    changed = 0
    try:
        doc = word.Documents.Open(str(docx_path))
        doc.Repaginate()

        for tbl in doc.Tables:
            if _looks_like_test11_3_table(tbl):
                _set_repeat_header(tbl, header_rows_count=header_rows_count)
                changed += 1

        if changed:
            doc.Save()
        doc.Close()
        return changed
    finally:
        word.Quit()


def update_all_fields_toc_headers(docx_path: str) -> None:
    """
    Полное обновление:
    - Repaginate
    - Fields.Update()
    - обновление полей в колонтитулах
    - обновление содержания (TOC), если есть
    """
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path))
        doc.Repaginate()

        # 1) Обновить поля в основном тексте
        try:
            doc.Fields.Update()
        except Exception:
            pass

        # 2) Колонтитулы
        for sec in doc.Sections:
            for hf_idx in (1, 2, 3):  # 1=Primary, 2=FirstPage, 3=EvenPages
                try:
                    sec.Headers(hf_idx).Range.Fields.Update()
                except Exception:
                    pass
                try:
                    sec.Footers(hf_idx).Range.Fields.Update()
                except Exception:
                    pass

        # 3) Содержание (TOC)
        try:
            if doc.TablesOfContents.Count >= 1:
                # Если TOC несколько — обновим все
                for i in range(1, doc.TablesOfContents.Count + 1):
                    try:
                        doc.TablesOfContents(i).Update()
                    except Exception:
                        pass
        except Exception:
            pass

        # 4) Ещё разок для пересчёта NUMPAGES после всех вставок/разрывов
        doc.Repaginate()
        try:
            doc.Fields.Update()
        except Exception:
            pass

        doc.Save()
        doc.Close()
    finally:
        word.Quit()


def finalize_docx(docx_path: str, *, test11_3_header_rows: int = 2) -> None:
    """
    Один вызов на финализацию документа:
    - сделать повтор шапки в таблице Тест 11.3 (на каждой странице)
    - обновить TOC/поля/колонтитулы (NUMPAGES/страницы в содержании)
    """
    enforce_test11_3_header_each_page(docx_path, header_rows_count=test11_3_header_rows)
    update_all_fields_toc_headers(docx_path)
