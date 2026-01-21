"""
word_table5_splitter.py

Разделение Таблицы 5 на 2 части + "Продолжение таблицы 5"
СТРОГО по фактическому переносу на новую страницу (по пагинации MS Word).

Гарантии:
- копируем ИМЕННО диапазон шапки (первые header_rows строк) -> НЕТ дублей
- шапка в продолжении ПОЛНАЯ (включая "Аттестационное испытание") за счет добора Range по Find()
- убираем пустой абзац СРАЗУ ПОСЛЕ ШАПКИ (между шапкой и продолжением) корректно
- НЕ даем "таблице-легенде" (Вероятность возникновения/Балл/...) прилипнуть к Таблице 5
- работает при vMerge, где Rows/Cell(row,col) могут ломаться

Требования:
- Windows
- установлен MS Word
- установлен pywin32 (win32com, pythoncom)
"""

from __future__ import annotations

import re
from typing import Optional

from logger import logger


# ------------------------------- helpers -------------------------------

def _norm_basic(s: str) -> str:
    s = "" if s is None else str(s)
    # Word добавляет маркеры конца ячейки/строки: \r и \x07
    s = s.replace("\r", " ").replace("\x07", " ")
    s = (
        s.replace("\u00A0", " ")
         .replace("\u202F", " ")
         .replace("\u200B", "")
         .replace("\u00AD", "")
    )
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s.replace("ё", "е")


def _looks_like_table5_by_placeholder(tbl) -> bool:
    """Самый надежный детектор — по плейсхолдерам."""
    try:
        txt = _norm_basic(tbl.Range.Text)
        return ("<<t5_risk>>" in txt) or ("t5_risk" in txt) or ("t5_" in txt)
    except Exception:
        return False


def _looks_like_table5_word(tbl) -> bool:
    """Fallback детектор по тексту таблицы."""
    try:
        txt = _norm_basic(tbl.Range.Text)
        must = (
            ("риск" in txt) and
            ("возможн" in txt and "причин" in txt) and
            ("аттестацион" in txt and "испыт" in txt)
        )
        placeholders = ("t5_" in txt)
        return must or placeholders
    except Exception:
        return False


def _find_table5_word(word_doc):
    """
    Ищем таблицу 5:
    0) по плейсхолдеру <<T5_RISK>> / T5_
    1) через Word Find по "Таблица 5" (берём таблицу после найденного места)
    2) через Word Find по "FMEA таблица 5"
    3) fallback: скан всех таблиц по тексту
    """
    caption_style = None

    # 0) placeholder scan
    try:
        for t in word_doc.Tables:
            if _looks_like_table5_by_placeholder(t):
                logger.info("Table5 split: нашли таблицу 5 по плейсхолдеру.")
                return t, None
    except Exception:
        pass

    def _try_find_and_take_table(find_text: str):
        nonlocal caption_style
        try:
            rng = word_doc.Content
            f = rng.Find
            f.ClearFormatting()
            ok = f.Execute(
                FindText=find_text,
                MatchCase=False,
                MatchWholeWord=False,
                MatchWildcards=False,
                Forward=True,
                Wrap=1,  # wdFindContinue
            )
            if not ok:
                return None

            try:
                caption_style = rng.Paragraphs(1).Range.Style
            except Exception:
                caption_style = None

            rng_after = word_doc.Range(rng.End, word_doc.Content.End)
            if int(rng_after.Tables.Count) > 0:
                t = rng_after.Tables(1)
                if _looks_like_table5_word(t):
                    return t
            return None
        except Exception:
            return None

    t = _try_find_and_take_table("Таблица 5")
    if t is not None:
        logger.info("Table5 split: нашли таблицу 5 через Find('Таблица 5').")
        return t, caption_style

    t = _try_find_and_take_table("FMEA таблица 5")
    if t is not None:
        logger.info("Table5 split: нашли таблицу 5 через Find('FMEA таблица 5').")
        return t, caption_style

    # fallback scan
    try:
        for tt in word_doc.Tables:
            if _looks_like_table5_word(tt):
                logger.info("Table5 split: нашли таблицу 5 сканированием всех таблиц.")
                return tt, caption_style
    except Exception:
        pass

    return None, caption_style


def _row_anchor_cell(t, row_idx: int):
    """
    Возвращает "якорную" ячейку строки, двигаясь справа налево.
    Нужна для случаев с вертикальными merge, когда некоторые Cell(row,col) недоступны.
    """
    try:
        cols = int(t.Columns.Count)
    except Exception:
        cols = 1

    for col in range(cols, 0, -1):
        try:
            return t.Cell(row_idx, col)
        except Exception:
            continue
    return None


def _get_first_table_after_pos(word_doc, pos: int):
    try:
        rng = word_doc.Range(pos, word_doc.Content.End)
        if int(rng.Tables.Count) > 0:
            return rng.Tables(1)
    except Exception:
        pass
    return None


def _copy_header_range_merge_safe(tbl, header_rows: int, c):
    """
    Делает Range, который покрывает первые header_rows строк таблицы.
    При vMerge MoveEnd(wdRow) может не включить вертикально-объединенную ячейку
    (например "Аттестационное испытание"), поэтому мы ДОБИРАЕМ конец range через Find().
    Возвращает Range (Duplicate) готовый к Copy().
    """
    hdr = tbl.Range.Duplicate
    hdr.Collapse(c.wdCollapseStart)

    moved = hdr.MoveEnd(Unit=c.wdRow, Count=int(header_rows))
    if int(moved) <= 0:
        raise RuntimeError("MoveEnd(wdRow) не смог сдвинуть конец диапазона шапки.")

    # добор по ключевым словам последней большой колонки шапки
    needles = ("Аттестацион", "испыт", "ПЧР")
    for needle in needles:
        try:
            fr = tbl.Range.Duplicate
            f = fr.Find
            f.ClearFormatting()
            ok = f.Execute(
                FindText=needle,
                MatchCase=False,
                MatchWholeWord=False,
                MatchWildcards=False,
                Forward=True,
                Wrap=0,  # wdFindStop
            )
            if not ok:
                continue

            cell = fr.Cells(1)

            # текст должен быть именно в шапке
            try:
                if int(cell.RowIndex) > int(header_rows):
                    continue
            except Exception:
                pass

            end_pos = int(cell.Range.End)
            if end_pos > int(hdr.End):
                hdr.End = end_pos
        except Exception:
            continue

    return hdr


def _looks_like_legend_table(tbl) -> bool:
    """
    Таблица-легенда после анализа рисков обычно содержит "Вероятность возникновения" и "Балл".
    Нужна, чтобы НЕ дать ей прилипнуть к таблице 5.
    """
    try:
        txt = _norm_basic(tbl.Range.Text)
        return ("вероятност" in txt) and ("балл" in txt) and ("уровен" in txt or "риска" in txt)
    except Exception:
        return False


def _ensure_paragraph_between_tables(word_doc, table_a, table_b, c) -> None:
    """
    Если две таблицы идут вплотную (Word может их склеить/перекинуть),
    гарантируем, что между ними есть хотя бы один абзац.
    """
    try:
        a_end = int(table_a.Range.End)
        b_start = int(table_b.Range.Start)
        if b_start <= a_end:
            return

        gap = word_doc.Range(a_end, b_start)
        t = gap.Text or ""

        # если между ними нет абзаца — вставим
        if "\r" not in t:
            r = word_doc.Range(a_end, a_end)
            r.InsertAfter("\r")
    except Exception:
        pass


# ------------------------------- main -------------------------------

def split_table5_with_continuation_open_doc(word_doc, *, header_rows: int = 2) -> bool:
    """
    Делит Таблицу 5 в ОТКРЫТОМ Word-документе (win32com), если она занимает > 1 страницы.
    Возвращает True, если было выполнено разделение, иначе False.
    """
    try:
        import win32com.client as win32
        c = win32.constants
    except Exception as e:
        logger.warning(f"Table5 split: pywin32 не доступен: {e}")
        return False

    logger.info(f"Table5 split: всего таблиц в документе: {int(word_doc.Tables.Count)}")

    tbl, caption_style = _find_table5_word(word_doc)
    if tbl is None:
        logger.warning("Table5 split: Таблица 5 не найдена — пропускаем.")
        return False

    # защита от повторного запуска
    try:
        after_text = word_doc.Range(tbl.Range.End, min(tbl.Range.End + 8000, word_doc.Content.End)).Text
        if re.search(r"продолжение\s+таблицы\s*5", after_text, flags=re.IGNORECASE):
            logger.info("Table5 split: продолжение уже есть — пропускаем.")
            return False
    except Exception:
        pass

    # Запретить Word разрезать одну строку пополам (если получится)
    try:
        tbl.Rows.AllowBreakAcrossPages = False
    except Exception:
        pass

    word_doc.Repaginate()

    def _page_at(pos: int) -> Optional[int]:
        try:
            r = word_doc.Range(pos, pos)
            return int(r.Information(c.wdActiveEndAdjustedPageNumber))
        except Exception:
            try:
                r = word_doc.Range(pos, pos)
                return int(r.Information(c.wdActiveEndPageNumber))
            except Exception:
                return None

    def _range_start_page(rng) -> Optional[int]:
        try:
            return _page_at(int(rng.Start))
        except Exception:
            return None

    def _range_end_page(rng) -> Optional[int]:
        try:
            end_pos = int(rng.End) - 1
            if end_pos < int(rng.Start):
                end_pos = int(rng.Start)
            return _page_at(end_pos)
        except Exception:
            return None

    start_page = _range_start_page(tbl.Range)
    end_page = _range_end_page(tbl.Range)

    if start_page is None or end_page is None:
        logger.warning("Table5 split: не удалось определить страницы таблицы.")
        return False

    if start_page == end_page:
        logger.info("Table5 split: таблица на одной странице — делить не нужно.")
        return False

    rows_count = int(tbl.Rows.Count)
    if rows_count <= header_rows + 1:
        logger.info("Table5 split: слишком мало строк — делить не нужно.")
        return False

    def _row_start_page(t, row_idx: int) -> Optional[int]:
        cell = _row_anchor_cell(t, row_idx)
        if not cell:
            return None
        return _page_at(int(cell.Range.Start))

    def _row_end_page(t, row_idx: int) -> Optional[int]:
        cell = _row_anchor_cell(t, row_idx)
        if not cell:
            return None
        end_pos = int(cell.Range.End) - 1
        if end_pos < int(cell.Range.Start):
            end_pos = int(cell.Range.Start)
        return _page_at(end_pos)

    # --- найти строку, которая уехала на следующую страницу (или ломается) ---
    split_row_idx: Optional[int] = None
    for r in range(header_rows + 1, rows_count + 1):
        rp_start = _row_start_page(tbl, r)
        if rp_start is None:
            continue
        if rp_start > start_page:
            split_row_idx = r
            break
        rp_end = _row_end_page(tbl, r)
        if rp_end is not None and rp_end > rp_start:
            split_row_idx = r
            break

    if split_row_idx is None:
        logger.warning("Table5 split: не нашли строку для переноса/разрыва — пропускаем.")
        return False

    logger.info(f"Table5 split: режем перед строкой {split_row_idx} (страницы {start_page} -> {end_page}).")

    sel = word_doc.Application.Selection

    # =====================================================================
    # 1) КОПИРУЕМ ТОЛЬКО ШАПКУ (Range первых header_rows строк) — без дублей
    # =====================================================================
    try:
        hdr_rng = _copy_header_range_merge_safe(tbl, header_rows, c)
        hdr_rng.Copy()
    except Exception as e:
        logger.warning(f"Table5 split: не удалось скопировать шапку диапазоном: {e}")
        return False

    # =====================================================================
    # 2) SplitTable основной таблицы (tbl) на две части
    # =====================================================================
    try:
        row_anchor = _row_anchor_cell(tbl, split_row_idx)
        if not row_anchor:
            logger.warning("Table5 split: не нашли якорную ячейку строки разреза.")
            return False

        sel.SetRange(int(row_anchor.Range.Start), int(row_anchor.Range.Start))
        sel.SplitTable()
    except Exception as e:
        logger.warning(f"Table5 split: SplitTable не сработал: {e}")
        return False

    # выйти из таблицы к разделяющему абзацу
    for _ in range(2000):
        try:
            if not sel.Information(c.wdWithInTable):
                break
            sel.MoveRight(Unit=c.wdCharacter, Count=1)
        except Exception:
            break

    # разрыв страницы
    try:
        sel.InsertBreak(Type=c.wdPageBreak)
    except Exception as e:
        logger.warning(f"Table5 split: не удалось вставить разрыв страницы: {e}")
        return False

    # "Продолжение таблицы 5"
    try:
        if caption_style is not None:
            try:
                sel.Style = caption_style
            except Exception:
                pass

        sel.ParagraphFormat.Alignment = c.wdAlignParagraphRight
        sel.ParagraphFormat.KeepWithNext = True

        sel.Font.Name = "Times New Roman"
        sel.Font.Size = 12
        sel.Font.Bold = True

        sel.TypeText("Продолжение таблицы 5")
        sel.TypeParagraph()
        sel.Font.Bold = False
    except Exception as e:
        logger.warning(f"Table5 split: не удалось вставить/оформить 'Продолжение таблицы 5': {e}")

    # =====================================================================
    # 3) Находим нижнюю таблицу (продолжение) и вставляем перед ней ШАПКУ
    # =====================================================================
    try:
        tbl2 = _get_first_table_after_pos(word_doc, int(sel.Range.End))
        if tbl2 is None:
            logger.warning("Table5 split: после 'Продолжение таблицы 5' не нашли таблицу.")
            return False
    except Exception as e:
        logger.warning(f"Table5 split: не удалось получить продолжение таблицы: {e}")
        return False

    try:
        # курсор на начало tbl2
        insert_pos = int(tbl2.Range.Start)
        sel.SetRange(insert_pos, insert_pos)
        sel.Collapse(c.wdCollapseStart)

        # выйти из tbl2 в абзац ПЕРЕД таблицей
        moved_out = False
        for _ in range(120):
            if not sel.Information(c.wdWithInTable):
                moved_out = True
                break
            sel.MoveLeft(Unit=c.wdCharacter, Count=1)

        if not moved_out:
            # если не смогли выйти — создадим абзац перед таблицей
            r0 = word_doc.Range(insert_pos, insert_pos)
            r0.InsertBefore("\r")
            sel.SetRange(insert_pos - 1, insert_pos - 1)

        # вставляем шапку (как таблицу)
        paste_pos = int(sel.Range.Start)
        sel.Paste()
        word_doc.Repaginate()

        # определяем вставленную таблицу-шапку
        header_tbl = None
        try:
            if sel.Information(c.wdWithInTable) and int(sel.Tables.Count) > 0:
                header_tbl = sel.Tables(1)
        except Exception:
            header_tbl = None

        if header_tbl is None:
            header_tbl = _get_first_table_after_pos(word_doc, paste_pos)

        if header_tbl is None:
            logger.warning("Table5 split: не удалось определить вставленную таблицу шапки.")
            return False

        # -----------------------------------------------------------------
        # КЛЮЧЕВОЕ: удалить пустой абзац СРАЗУ ПОСЛЕ ШАПКИ
        # Делаем это НАДЕЖНО: берем таблицу, которая идет СРАЗУ ПОСЛЕ header_tbl,
        # и удаляем промежуток ТОЛЬКО между ними.
        # -----------------------------------------------------------------
        data_tbl = _get_first_table_after_pos(word_doc, int(header_tbl.Range.End))
        if data_tbl is None:
            logger.warning("Table5 split: не нашли таблицу продолжения после вставленной шапки.")
            return False

        # если вдруг data_tbl == header_tbl (редко, но бывает) — сдвинем позицию на 1 символ
        try:
            if int(data_tbl.Range.Start) == int(header_tbl.Range.Start):
                data_tbl = _get_first_table_after_pos(word_doc, int(header_tbl.Range.End) + 1)
        except Exception:
            pass

        if data_tbl is None:
            logger.warning("Table5 split: не удалось получить таблицу данных для склейки.")
            return False

        # удаляем РОВНО разрыв между header_tbl и data_tbl -> склеиваются
        try:
            join_rng = word_doc.Range(int(header_tbl.Range.End), int(data_tbl.Range.Start))
            if int(join_rng.End) > int(join_rng.Start):
                join_rng.Delete()
        except Exception as e:
            logger.warning(f"Table5 split: не удалось удалить промежуток между шапкой и продолжением: {e}")

        # после склейки "таблица 5 (продолжение)" — это header_tbl (как первая)
        cont_tbl = header_tbl

        # делаем шапку повторяющейся (если получится)
        try:
            for rr in range(1, header_rows + 1):
                try:
                    cont_tbl.Rows(rr).HeadingFormat = True
                except Exception:
                    pass
        except Exception:
            pass

        # -----------------------------------------------------------------
        # ВАЖНО: НЕ ДАЕМ СЛЕДУЮЩЕЙ ТАБЛИЦЕ (легенде) ПРИЛИПНУТЬ К ТАБЛИЦЕ 5
        # Если сразу после cont_tbl идет таблица-легенда, гарантируем абзац между ними.
        # -----------------------------------------------------------------
        try:
            next_tbl = _get_first_table_after_pos(word_doc, int(cont_tbl.Range.End))
            if next_tbl is not None and _looks_like_legend_table(next_tbl):
                _ensure_paragraph_between_tables(word_doc, cont_tbl, next_tbl, c)
        except Exception:
            pass

        try:
            logger.info(
                f"Table5 split: header inserted+merged; header_rows={header_rows}; "
                f"cols={int(cont_tbl.Columns.Count)}"
            )
        except Exception:
            pass

    except Exception as e:
        logger.warning(f"Table5 split: не удалось вставить/склеить шапку во 2-ю часть: {e}")
        return False

    logger.info("Table5 split: выполнено.")
    return True


def update_fields_with_word(docx_path: str) -> None:
    """
    Открывает DOCX в Word, режет Таблицу 5, обновляет поля/колонтитулы, сохраняет.
    ВАЖНО: вызывается из QThread => обязательно CoInitialize/CoUninitialize.
    """
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as e:
        logger.warning(f"Word postprocess пропущен: pywin32 не установлен/недоступен: {e}")
        return

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(str(docx_path), ReadOnly=False, AddToRecentFiles=False)

        try:
            doc.Repaginate()
            split_table5_with_continuation_open_doc(doc, header_rows=2)
            doc.Repaginate()
        except Exception as e:
            logger.warning(f"Table5 split: ошибка во время разрезания: {e}")

        # обновление полей
        try:
            doc.Fields.Update()
        except Exception:
            pass

        try:
            for sec in doc.Sections:
                sec.Headers(1).Range.Fields.Update()
                sec.Footers(1).Range.Fields.Update()
        except Exception:
            pass

        doc.Save()
        doc.Close(SaveChanges=False)
        doc = None

    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
