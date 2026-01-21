# io_manager.py
from __future__ import annotations

from pathlib import Path
from typing import List, Dict

import pandas as pd
from logger import logger


# ---------------------------
# Общие утилиты
# ---------------------------

def validate_file(path: Path) -> None:
    """Проверка существования файла на диске."""
    if not path.exists():
        logger.error(f"Файл не найден: {path}")
        raise FileNotFoundError(f"Файл не найден: {path}")


# ---------------------------
# Тесты (Excel -> список строк)
# ---------------------------

def load_tests_list(excel_path: Path) -> List[str]:
    """
    Читает Excel с тестами и возвращает список названий тестов.
    Берётся первый столбец первого листа (как было ранее).
    """
    validate_file(excel_path)
    df = pd.read_excel(excel_path, dtype=str)

    if df.empty or df.columns.size == 0:
        raise ValueError(f"В файле {excel_path} нет данных для тестов")

    tests = (
        df.iloc[:, 0]
        .dropna()
        .astype(str)
        .map(str.strip)
        .tolist()
    )
    logger.debug(f"Загружено тестов: {len(tests)} из {excel_path.name}")
    return tests


# ---------------------------
# Оборудование (Excel -> списки словарей)
# ---------------------------

def _parse_equipment_df(df: pd.DataFrame) -> List[dict]:
    """
    Универсальный парсер листа Excel с оборудованием.
    Возвращает список словарей с ключами: name_sn, params, cert, date, until
    """
    df = df.copy()
    df.rename(columns=lambda c: str(c).strip(), inplace=True)

    name_col  = next((c for c in df.columns if "Наименование" in c), None)
    sn_col    = next((c for c in df.columns if "зав" in str(c).lower()), None)
    param_col = next((c for c in df.columns if "Определяемые показатели" in c), None)
    cert_col  = next((c for c in df.columns if "№ свидетельств" in c or "№ свидетельтва" in c), None)
    date_col  = next((c for c in df.columns if "Дата поверки" in c), None)
    until_col = next((c for c in df.columns if "Срок действия поверки" in c), None)

    if not name_col or not sn_col:
        raise ValueError("Нет колонок «Наименование»/«Зав. (серийный) номер»")

    items: List[dict] = []
    for _, row in df.iterrows():
        nm = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
        if not nm:
            continue
        sn = str(row[sn_col]).strip() if pd.notna(row[sn_col]) else ""
        items.append({
            "name_sn": f"{nm}, {sn}".strip(", "),
            "params": str(row[param_col]).strip() if param_col and pd.notna(row.get(param_col)) else "",
            "cert":   str(row[cert_col]).strip()  if cert_col  and pd.notna(row.get(cert_col))  else "",
            "date":   str(row[date_col]).strip()  if date_col  and pd.notna(row.get(date_col))  else "",
            "until":  str(row[until_col]).strip() if until_col and pd.notna(row.get(until_col)) else "",
        })
    return items


def load_equipment_by_sheets(excel_path: Path) -> Dict[str, List[dict]]:
    """
    Возвращает {имя_листа: [элементы...]}. Все листы обрабатываются одинаково.
    Листы с ошибками парсинга не валят процесс — для них будет [] и warning в лог.
    """
    validate_file(excel_path)
    xls = pd.ExcelFile(excel_path)
    if not xls.sheet_names:
        raise ValueError(f"В файле {excel_path} нет листов Excel")

    result: Dict[str, List[dict]] = {}
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
            result[sheet] = _parse_equipment_df(df)
            logger.debug(f"[{sheet}] загружено {len(result[sheet])} позиций")
        except Exception as e:
            logger.warning(f"Лист «{sheet}» пропущен: {e}")
            result[sheet] = []
    return result


def load_equipment_list(excel_path: Path) -> List[dict]:
    """
    Обратная совместимость: берёт первый лист (как в старой версии).
    """
    validate_file(excel_path)
    df = pd.read_excel(excel_path, dtype=str)
    return _parse_equipment_df(df)
