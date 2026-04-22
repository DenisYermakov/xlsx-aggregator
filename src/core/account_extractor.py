"""Извлечение троек счетов AB/AN/AS и BB/BN/BS по маскам полей на листах PF."""

from __future__ import annotations

import re
from typing import Literal

from src.core.config import ACCOUNT_SOURCE_SHEETS, SHEET_FIELD_PREFIX


def _norm(v: object) -> str:
    return "" if v is None else str(v).strip()


def extract_account_slots(
    headers: list[str | None],
    sheet_prefix: str,
) -> list[tuple[str, Literal["A", "B"], tuple[str, str, str]]]:
    """Возвращает список слотов, где на листе есть полные тройки AB/AN/AS или BB/BN/BS.

    ТЗ: «Ключ счета равен маске поля AB*,AN*,AS* … поля идут последовательно …»

    Args:
        headers: Имена колонок со строки 2 (могут содержать ``None``).
        sheet_prefix: Префикс листа: ``'YW2'`` | ``'YW3'`` | ``'YWJ1'``.

    Returns:
        Список элементов ``(slot_id, "A"|"B", (col_AB, col_AN, col_AS))`` для группы A
        или аналогично для B с ``BB, BN, BS``.
    """
    slots: dict[str, dict[str, str]] = {}
    pat = re.compile(rf"^{re.escape(sheet_prefix)}(AB|AN|AS|BB|BN|BS)(.+)$")
    for h in headers:
        if h is None:
            continue
        m = pat.match(str(h))
        if not m:
            continue
        kind, slot = m.group(1), m.group(2)
        slots.setdefault(slot, {})[kind] = str(h)

    result: list[tuple[str, Literal["A", "B"], tuple[str, str, str]]] = []
    for slot, kinds in slots.items():
        if {"AB", "AN", "AS"} <= kinds.keys():
            result.append(
                (slot, "A", (kinds["AB"], kinds["AN"], kinds["AS"]))
            )
        if {"BB", "BN", "BS"} <= kinds.keys():
            result.append(
                (slot, "B", (kinds["BB"], kinds["BN"], kinds["BS"]))
            )
    return result


def extract_all_accounts(
    sheets: dict[str, list[dict[str, object]]],
) -> set[tuple[str, str, str]]:
    """Извлекает уникальные ключи счетов со всех sheet-источников.

    ТЗ: «На Листах YW2PF, YW3PF и YWJ1PF указаны счета. Ключ счета равен
    маске поля AB*,AN*,AS* или BB*,BN*,BS*.»

    Алгоритм:
        1. Для каждого листа (YW2PF, YW3PF, YWJ1PF) по префиксу
           (YW2/YW3/YWJ1) находим слоты с полными тройками.
        2. Берём только слоты, где присутствуют ВСЕ ТРИ поля тройки.
        3. Группы A и B сливаются в один set троек для JOIN с SCAB/SCAN/SCAS.
        4. Пустые тройки (все три компонента = "") не добавляем.

    Args:
        sheets: ``{имя_листа: [row_dict, ...]}`` — как из ``read_sheet_as_dicts``.

    Returns:
        Множество уникальных троек ``(AB, AN, AS)`` как строк.
    """
    accounts: set[tuple[str, str, str]] = set()
    for sheet_name in ACCOUNT_SOURCE_SHEETS:
        prefix_key = sheet_name
        if prefix_key not in SHEET_FIELD_PREFIX:
            continue
        sheet_prefix = SHEET_FIELD_PREFIX[prefix_key]
        rows = sheets.get(sheet_name) or []
        if not rows:
            continue
        headers = list(rows[0].keys())
        slots = extract_account_slots(headers, sheet_prefix)
        for row in rows:
            for _slot_id, _group, (c1, c2, c3) in slots:
                v1, v2, v3 = row.get(c1), row.get(c2), row.get(c3)
                triple = (_norm(v1), _norm(v2), _norm(v3))
                if any(triple):
                    accounts.add(triple)
    return accounts
