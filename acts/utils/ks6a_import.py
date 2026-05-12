"""
Импорт строк работ из журнала КС-6а (Excel).

Ищет строку заголовков по типичным подписям граф формы КС-6а и сопоставляет столбцы.
Подходит для типовых шаблонов с графами: наименование, ед. изм., количество, цена, суммы.
"""
from __future__ import annotations

import logging
import re
from decimal import Decimal, InvalidOperation
from typing import Any

import aspose.cells as ac

logger = logging.getLogger(__name__)

MAX_PARSE_ROWS = 500
HEADER_SCAN_ROWS = 90
HEADER_SCAN_COLS = 40
DATA_SCAN_ROWS = 400


def _hyphen_fold(s: str) -> str:
    """Склеивает переносы через дефис, типичные для печатных форм в Excel (еди- ничной)."""
    t = str(s).replace("\n", " ").replace("\r", " ")
    t = re.sub(r"-\s+", "", t)
    return t


def _norm(s: Any) -> str:
    if s is None:
        return ""
    t = _hyphen_fold(str(s))
    t = t.lower().replace("ё", "е")
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _cell_str(ws: Any, row_0: int, col_0: int) -> str:
    try:
        v = ws.cells.get(row_0, col_0).value
    except Exception:
        return ""
    if v is None:
        return ""
    return str(v).strip()


def _to_decimal(val: Any) -> Decimal | None:
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        try:
            return Decimal(str(val))
        except InvalidOperation:
            return None
    s = str(val).strip().replace("\xa0", "").replace(" ", "")
    s = s.replace(",", ".").replace("−", "-")
    if not s or s == "-":
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


# (поле внутренней карты, ключевые фразы в заголовке, чем длиннее — тем приоритетнее)
HEADER_RULES: list[tuple[str, tuple[str, ...]]] = [
    (
        "name",
        (
            "конструктивные элементы и виды работ",
            "конструктивные элементы",
            "наименование работ и затрат",
            "наименование работ",
            "наименование",
        ),
    ),
    ("position", ("шифр и номер позиции", "номер позиции по смете", "позиции по смете", "позиция по смете")),
    ("code", ("код по приказу", "шифр", "код")),
    ("number_pricelist", ("номер единичной расценки", "единичной расценки", "номер расценки", "№ расценки")),
    ("unit", ("единица измерения", "ед. изм", "ед изм")),
    ("quantity", ("количество работ по смете", "количество", "объем", "объём", "выполнено")),
    ("price", ("цена за единицу", "цена за ед", "цена, руб", "цена за единицу работ", "цена")),
    ("cost_from_start", ("с начала проведения работ", "с начала работ", "с начала проведения")),
    ("cost_from_year", ("с начала года", "с начала текущего года")),
    ("cost_report_period", ("за отчетный период", "за отчётный период", "за период", "за месяц")),
    (
        "cost_total",
        (
            "сметная (договорная) стоимость",
            "сметная стоимость",
            "договорная стоимость",
            "всего затрат",
            "стоимость всего",
            "фактическая стоимость",
            "сумма",
        ),
    ),
]


# Строки в графе наименования, которые не являются видами работ
_SKIP_NAME_NORMALIZED = frozenset(
    {
        "конструктивные элементы и виды работ",
        "наименование работ и затрат",
        "наименование работ",
        "наименование",
    }
)


def _merged_header_cell_text(ws: Any, r: int, c: int) -> str:
    """Текст заголовка из строки r и r+1 (двухстрочные шапки КС-6а)."""
    parts: list[str] = []
    for rr in (r, r + 1):
        if rr >= HEADER_SCAN_ROWS:
            break
        s = _cell_str(ws, rr, c)
        if s:
            parts.append(s)
    return _norm(" ".join(parts))


def _best_header_row(ws: Any) -> tuple[int | None, dict[str, int]]:
    """Возвращает (индекс строки заголовка 0-based, сопоставление поле -> индекс столбца)."""
    best_row: int | None = None
    best_score = 0
    best_map: dict[str, int] = {}

    for r in range(HEADER_SCAN_ROWS - 1):
        row_texts = [_merged_header_cell_text(ws, r, c) for c in range(HEADER_SCAN_COLS)]
        colmap: dict[str, int] = {}
        score = 0
        for field, phrases in HEADER_RULES:
            best_c = None
            best_phrase_len = 0
            for c, txt in enumerate(row_texts):
                if not txt:
                    continue
                for ph in sorted(phrases, key=len, reverse=True):
                    if ph in txt and len(ph) > best_phrase_len:
                        best_c = c
                        best_phrase_len = len(ph)
                        break
            if best_c is not None:
                colmap[field] = best_c
                score += best_phrase_len + 2
        if "name" in colmap and score > best_score:
            best_score = score
            best_row = r
            best_map = colmap

    return best_row, best_map


def _decimal_for_session(d: Decimal | None) -> str:
    if d is None:
        return ""
    return format(d, "f")


def parse_ks6a_file(path: str) -> list[dict[str, Any]]:
    """
    Читает первый лист книги Excel и возвращает до MAX_PARSE_ROWS словарей
    для initial формсета (строки для КС-2 и КС-3 одновременно).

    Числовые значения — строки, пригодные для JSON-сессии и DecimalField.
    """
    wb = None
    try:
        wb = ac.Workbook(path)
        ws = wb.worksheets[0]
        header_row, colmap = _best_header_row(ws)
        if header_row is None or "name" not in colmap:
            logger.warning("КС-6а: не найдена строка заголовков с графой «наименование»")
            return []

        name_c = colmap["name"]
        out: list[dict[str, Any]] = []
        empty_streak = 0

        for r in range(header_row + 1, header_row + 1 + DATA_SCAN_ROWS):
            name = _cell_str(ws, r, name_c)
            name_n = _norm(name)
            if not name or name in ("-", "—", "–"):
                empty_streak += 1
                if empty_streak >= 5 and out:
                    break
                continue
            empty_streak = 0

            if name_n in _SKIP_NAME_NORMALIZED:
                continue
            if len(name) < 3:
                continue
            if name.strip().isdigit() and len(name.strip()) <= 2:
                continue

            def col(field: str) -> str:
                if field not in colmap:
                    return ""
                return _cell_str(ws, r, colmap[field])

            position = col("position") if "position" in colmap else ""
            code = col("code") if "code" in colmap else ""
            if not position:
                position = code
            if not code:
                code = position
            number_pricelist = col("number_pricelist")
            unit = col("unit")

            qty = _to_decimal(ws.cells.get(r, colmap["quantity"]).value) if "quantity" in colmap else None
            price = _to_decimal(ws.cells.get(r, colmap["price"]).value) if "price" in colmap else None

            cfs = _to_decimal(ws.cells.get(r, colmap["cost_from_start"]).value) if "cost_from_start" in colmap else None
            cfy = _to_decimal(ws.cells.get(r, colmap["cost_from_year"]).value) if "cost_from_year" in colmap else None
            crp = _to_decimal(ws.cells.get(r, colmap["cost_report_period"]).value) if "cost_report_period" in colmap else None
            ctot = _to_decimal(ws.cells.get(r, colmap["cost_total"]).value) if "cost_total" in colmap else None

            if qty is None and price is None and crp is None and ctot is None:
                # нет чисел — пропускаем подписи/пустые блоки
                continue

            line_total = None
            if qty is not None and price is not None:
                line_total = (qty * price).quantize(Decimal("0.01"))
            if crp is None and line_total is not None:
                crp = line_total
            if crp is None and ctot is not None:
                crp = ctot
            if cfs is None:
                cfs = Decimal("0")
            if cfy is None:
                cfy = Decimal("0")
            if crp is None:
                crp = Decimal("0")

            # Для КС-2: только суммы за период — трактуем как 1 × сумма
            if qty is None and price is None and crp > 0:
                qty = Decimal("1")
                price = crp

            order = len(out)
            row_dict: dict[str, Any] = {
                "position": position,
                "name": name,
                "code": code or position,
                "number_pricelist": number_pricelist,
                "unit": unit,
                "quantity": _decimal_for_session(qty) if qty is not None else "",
                "price": _decimal_for_session(price) if price is not None else "",
                "cost_from_start": _decimal_for_session(cfs),
                "cost_from_year": _decimal_for_session(cfy),
                "cost_report_period": _decimal_for_session(crp),
                "order": order,
            }
            out.append(row_dict)
            if len(out) >= MAX_PARSE_ROWS:
                break

        return out
    finally:
        if wb is not None:
            del wb
