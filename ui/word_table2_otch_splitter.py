from __future__ import annotations

import re
from typing import Optional

import win32com.client as win32
from win32com.client import constants as c


def _clean_cell_text(s: str) -> str:
    # Word возвращает текст ячейки с '\r\x07'
    s = (s or "").replace("\r", "").replace("\x07", "")
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _norm(s: str) -> str:
    return _clean_cell_text(s).lower().replace("ё", "е")


def _looks_like_table2(word_table) -> bool:
    """
    Узнаём Таблицу 2 по шапке:
    'Тест', 'Критерий приемлемости', 'Фактические результаты', 'Оценка'
    """
    try:
        if word_table.Rows.Count < 1:
            return False
        row1 = word_table.Rows(1)
        head = " ".join(_clean_cell_text(cell.Range.Text) for cell in row1.Cells)
        h = _norm(head)
        return ("тест" in h and "критер" in h and "фактичес" in h and "оценк" in h)
    except Exception:
        return False


def _insert_continuation_block_before_table(word_app, doc, table, title: str) -> None:
    """
    Вставляет ПЕРЕД table:
      - абзац 'Продолжение таблицы 2' (TNR12 bold, справа)
      - абзац начинается С НОВОЙ СТРАНИЦЫ (PageBreakBefore=True)
    ВАЖНО: гарантированно вставляем ВНЕ таблицы (не в ячейку).
    """
    sel = word_app.Selection

    # Ставим курсор в начало новой таблицы
    sel.SetRange(table.Range.Start, table.Range.Start)

    # Гарантированно выходим из таблицы в "параграф-якорь" перед ней
    # Иногда Word держит Range внутри первой ячейки — выходим MoveLeft'ом.
    guard = 0
    while sel.Information(c.wdWithInTable) and guard < 50:
        sel.MoveLeft(Unit=c.wdCharacter, Count=1)
        guard += 1

    # Теперь мы вне таблицы: вставляем абзац перед таблицей
    sel.InsertParagraphAfter()          # создаём параграф перед таблицей
    sel.MoveUp(Unit=c.wdParagraph, Count=1)

    # Настраиваем абзац как "начало новой страницы"
    sel.ParagraphFormat.PageBreakBefore = True
    sel.ParagraphFormat.Alignment = c.wdAlignParagraphRight
    sel.ParagraphFormat.SpaceBefore = 0
    sel.ParagraphFormat.SpaceAfter = 0

    # Текст + шрифт
    sel.Font.Name = "Times New Roman"
    sel.Font.Size = 12
    sel.Font.Bold = True
    sel.TypeText(title)

    # Важно: НЕ InsertBreak() — иначе может появляться пустая страница
    # И НЕ InsertParagraphAfter() — иначе появится лишняя пустая строка


def split_table2_with_continuation_word(
    docx_path: str,
    *,
    continuation_title: str = "Продолжение таблицы 2",
    header_rows_count: int = 1,
    max_splits: int = 20,
) -> bool:
    """
    Режет Таблицу 2 по фактическому переносу на следующую страницу (через Word пагинацию).
    Возвращает True если хотя бы один разрез сделан.
    """
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False

    did_any = False
    try:
        doc = word.Documents.Open(str(docx_path))

        # если уже есть "Продолжение таблицы 2" — не плодим дубли
        if continuation_title in _clean_cell_text(doc.Content.Text):
            doc.Close(SaveChanges=False)
            return False

        doc.Repaginate()

        # найдём таблицу 2
        t2 = None
        for t in doc.Tables:
            if _looks_like_table2(t):
                t2 = t
                break

        if t2 is None:
            doc.Close(SaveChanges=False)
            return False

        # сделаем шапку повторяемой
        try:
            for i in range(1, header_rows_count + 1):
                t2.Rows(i).HeadingFormat = True
        except Exception:
            pass

        splits_done = 0
        cur_table = t2

        while splits_done < max_splits:
            doc.Repaginate()

            # страница, на которой заканчивается шапка
            base_page = cur_table.Rows(header_rows_count).Range.Information(c.wdActiveEndPageNumber)

            # найдём первую строку данных, которая ушла на следующую страницу
            split_row_idx: Optional[int] = None
            for ri in range(header_rows_count + 1, cur_table.Rows.Count + 1):
                pg = cur_table.Rows(ri).Range.Information(c.wdActiveEndPageNumber)
                if pg > base_page:
                    split_row_idx = ri
                    break

            if split_row_idx is None:
                break  # таблица целиком помещается на странице

            # 1) Копируем текущую таблицу после неё же
            count_before = doc.Tables.Count
            ins = cur_table.Range
            ins.Collapse(c.wdCollapseEnd)
            ins.InsertParagraphAfter()
            ins.Collapse(c.wdCollapseEnd)

            cur_table.Range.Copy()
            ins.PasteAndFormat(c.wdFormatOriginalFormatting)

            new_table = doc.Tables(count_before + 1)

            # 2) В НОВОЙ таблице удаляем строки ДО split_row_idx (кроме шапки)
            for _ in range((split_row_idx - 1) - header_rows_count):
                new_table.Rows(header_rows_count + 1).Delete()

            # 3) В СТАРОЙ таблице удаляем строки split_row_idx..конец
            while cur_table.Rows.Count >= split_row_idx:
                cur_table.Rows(split_row_idx).Delete()

            # 4) Вставляем "Продолжение таблицы 2" ПЕРЕД новой таблицей (вне таблицы!)
            _insert_continuation_block_before_table(word, doc, new_table, continuation_title)

            # 5) Повторяемая шапка для новой таблицы
            try:
                for i in range(1, header_rows_count + 1):
                    new_table.Rows(i).HeadingFormat = True
            except Exception:
                pass

            did_any = True
            splits_done += 1
            cur_table = new_table

        doc.Save()
        doc.Close()
        return did_any

    finally:
        word.Quit()


def update_fields_and_split_table2(
    docx_path: str,
    *,
    continuation_title: str = "Продолжение таблицы 2",
    header_rows_count: int = 1,
) -> bool:
    """
    1) Режет Таблицу 2 по фактическому переносу страницы (через Word).
    2) Обновляет поля документа + колонтитулы.
    Возвращает True если были сделаны разрезы.
    """
    did_split = split_table2_with_continuation_word(
        docx_path,
        continuation_title=continuation_title,
        header_rows_count=header_rows_count,
    )

    # Обновление полей/колонтитулов отдельным проходом (надежно)
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path))
        doc.Repaginate()

        doc.Fields.Update()
        for sec in doc.Sections:
            try:
                sec.Headers(1).Range.Fields.Update()
            except Exception:
                pass
            try:
                sec.Footers(1).Range.Fields.Update()
            except Exception:
                pass

        doc.Save()
        doc.Close()
    finally:
        word.Quit()

    return did_split
