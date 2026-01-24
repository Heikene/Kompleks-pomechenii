# word_update_all.py
from __future__ import annotations

import win32com.client as win32
from win32com.client import constants as c


def update_all_fields_toc_and_pagination(docx_path: str) -> None:
    """
    1) Репагинация (чтобы NUMPAGES стало корректным)
    2) Обновление оглавления (если оно настоящее)
    3) Обновление всех полей + полей в колонтитулах
    """
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path))

        # Важно: несколько проходов репагинации реально помогают после больших вставок/разрезаний
        doc.Repaginate()

        # Оглавление (если в документе есть TOC-поле)
        try:
            # обновить все оглавления
            if doc.TablesOfContents.Count > 0:
                for i in range(1, doc.TablesOfContents.Count + 1):
                    doc.TablesOfContents(i).Update()
        except Exception:
            pass

        # Все поля документа
        doc.Fields.Update()

        # Поля в колонтитулах/футерах (в т.ч. PAGE/NUMPAGES)
        for sec in doc.Sections:
            try:
                sec.Headers(c.wdHeaderFooterPrimary).Range.Fields.Update()
            except Exception:
                pass
            try:
                sec.Footers(c.wdHeaderFooterPrimary).Range.Fields.Update()
            except Exception:
                pass

        doc.Repaginate()
        doc.Save()
        doc.Close()
    finally:
        word.Quit()
