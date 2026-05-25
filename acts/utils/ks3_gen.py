"""
Генератор справки КС-3.
Использует Aspose.Cells для заполнения шаблона и экспорта в PDF/XLSX.
"""
import gc
import logging
import os
import traceback

import aspose.cells as ac
from aspose.cells import SaveFormat

try:
    from django.conf import settings
    DJANGO_AVAILABLE = True
except ImportError:
    DJANGO_AVAILABLE = False

logger = logging.getLogger(__name__)


# -------------------------------------------------------------------
# Вспомогательная функция для установки числа с заданным форматом
# -------------------------------------------------------------------
def set_cell_number_format(ws, row0, col0, value, number_format):
    """
    Устанавливает значение ячейки и числовой формат.
    ws   – объект Worksheet
    row0 – индекс строки (0-based)
    col0 – индекс столбца (0-based)
    value – числовое значение
    number_format – строка формата, например "#,##0.00"
    """
    cell = ws.cells.get(row0, col0)
    cell.put_value(value)
    style = cell.get_style()
    style.custom = number_format
    cell.set_style(style)


# -------------------------------------------------------------------
# Константы шаблона КС-3 (номера строк — 1-based)
# -------------------------------------------------------------------
KS3_PAGES = [
    {"page": 1, "rows": list(range(31, 58))},
    {"page": 2, "rows": list(range(68, 115))},
    {"page": 3, "rows": list(range(129, 163))},
]
PAGE1_FILL_WORK_COUNT = 25
PAGE1_LAST_WORK_ROW = 57
PAGE2_FIRST_WORK_ROW = 68
PAGE2_LAST_WORK_ROW = 114
PAGE2_HEADER_START = 58
PAGE2_HEADER_END = 67
PAGE3_FIRST_WORK_ROW = 129
PAGE3_LAST_WORK_ROW = 162
PAGE3_HEADER_START = 115
PAGE3_HEADER_END = 128
PAGE2_PAGE_BREAK_ROW = 60
PAGE2_PAGE_BREAK_AFTER_HEADER = 61
PAGE3_PAGE_BREAK_ROW = 119
PAGE3_PAGE_BREAK_AFTER_HEADER = 120
PAGE2_FILL_WORK_COUNT = 46
WORKS_BEFORE_PAGE3 = PAGE1_FILL_WORK_COUNT + PAGE2_FILL_WORK_COUNT
PAGE3_WORK_COUNT = len(KS3_PAGES[2]["rows"])
MAX_WORK_COUNT = WORKS_BEFORE_PAGE3 + PAGE3_WORK_COUNT
TOTAL_BLOCK_START_ROW = 164  # 163 — пустая строка-разделитель в шаблоне
SIGNATURE_BLOCK_START_ROW = 168
SIGNATURE_BLOCK_END_ROW = 177
SIGNATURE_ON_PAGE1_MAX_WORKS = 13
SIGNATURE_PAGE2_LAYOUT_MIN_WORKS = 26
SIGNATURE_ON_PAGE2_MAX_WORKS = 61
SIGNATURE_PAGE2_ONLY_NEXT_MIN_WORKS = 62
SIGNATURE_PAGE2_ONLY_NEXT_MAX_WORKS = 71
SIGNATURE_PAGE_SAFETY_MARGIN = 25

COL_NUM = 0
COL_NAME = 3
COL_CODE = 26
COL_COST_START = 30
COL_COST_YEAR = 37
COL_COST_PERIOD = 45
TOTAL_COL = 45
LABEL_COL = 44

TOTAL_LABELS = {
    "total": "Итого",
    "vat": "Сумма НДС",
    "total_with_vat": "Всего с учетом НДС",
}


# -------------------------------------------------------------------
# Разрывы страниц и сжатие листа
# -------------------------------------------------------------------
def _row_height_pt(worksheet, row0):
    height = worksheet.cells.get_row_height(row0)
    if height <= 0:
        height = worksheet.standard_height
    return float(height)


def _printable_page_height_pt(worksheet):
    page_setup = worksheet.page_setup
    top = float(page_setup.top_margin) * 72
    bottom = float(page_setup.bottom_margin) * 72
    paper_height = 842.0
    zoom = float(page_setup.zoom) / 100.0 if page_setup.is_percent_scale else 1.0
    return (paper_height - top - bottom) * zoom


def _manual_section_start(row0, manual_breaks):
    section_start = 0
    for break_row in manual_breaks:
        if row0 >= break_row:
            section_start = break_row
        else:
            break
    return section_start


def _print_page_in_section(worksheet, row_1based, page_height, manual_breaks):
    row0 = row_1based - 1
    section_start = _manual_section_start(row0, manual_breaks)
    page_index = 0
    used_height = 0.0
    for r in range(section_start, row0 + 1):
        height = _row_height_pt(worksheet, r)
        if used_height > 0 and used_height + height > page_height:
            page_index += 1
            used_height = 0.0
        used_height += height
    return page_index


def _rows_height_pt(worksheet, start_row_1based, end_row_1based):
    total = 0.0
    for row0 in range(start_row_1based - 1, end_row_1based):
        total += _row_height_pt(worksheet, row0)
    return total


def _remaining_height_before_row(worksheet, row_1based, manual_breaks, page_height):
    row0 = row_1based - 1
    section_start = _manual_section_start(row0, manual_breaks)
    used_height = 0.0
    for r in range(section_start, row0):
        height = _row_height_pt(worksheet, r)
        if used_height > 0 and used_height + height > page_height:
            used_height = 0.0
        used_height += height
    return page_height - used_height


def _add_page_break_before_row(worksheet, row_1based):
    break_row = row_1based - 1
    for existing in worksheet.horizontal_page_breaks:
        if existing.row == break_row:
            return False
    worksheet.horizontal_page_breaks.add(break_row)
    return True


def _remove_page_break_before_row(worksheet, row_1based):
    break_row = row_1based - 1
    breaks = worksheet.horizontal_page_breaks
    to_remove = [i for i, br in enumerate(breaks) if br.row == break_row]
    for i in reversed(to_remove):
        breaks.remove_at(i)
    return bool(to_remove)


def _row_block_is_split(worksheet, block_start, block_end):
    manual_breaks = sorted(br.row for br in worksheet.horizontal_page_breaks)
    page_height = _printable_page_height_pt(worksheet)
    pages = {
        _print_page_in_section(worksheet, row, page_height, manual_breaks)
        for row in range(block_start, block_end + 1)
    }
    return len(pages) > 1


def _get_all_work_rows():
    rows = list(KS3_PAGES[0]["rows"][:PAGE1_FILL_WORK_COUNT])
    rows.extend(KS3_PAGES[1]["rows"][:PAGE2_FILL_WORK_COUNT])
    rows.extend(KS3_PAGES[2]["rows"])
    return rows


def _uses_page2(works_count):
    return works_count > PAGE1_FILL_WORK_COUNT


def _uses_page3(works_count):
    return works_count > WORKS_BEFORE_PAGE3


def _signature_can_stay_with_works(works_count):
    if works_count <= SIGNATURE_ON_PAGE1_MAX_WORKS:
        return True
    if (
        SIGNATURE_PAGE2_LAYOUT_MIN_WORKS
        <= works_count
        <= SIGNATURE_ON_PAGE2_MAX_WORKS
        and not _uses_page3(works_count)
    ):
        return True
    return False


def _signature_must_be_on_next_page(works_count):
    if (
        SIGNATURE_ON_PAGE1_MAX_WORKS
        < works_count
        < SIGNATURE_PAGE2_LAYOUT_MIN_WORKS
    ):
        return True
    if (
        not _uses_page3(works_count)
        and SIGNATURE_PAGE2_ONLY_NEXT_MIN_WORKS
        <= works_count
        <= SIGNATURE_PAGE2_ONLY_NEXT_MAX_WORKS
    ):
        return True
    return False


def _setup_page2_page_break(worksheet):
    _remove_page_break_before_row(worksheet, PAGE2_PAGE_BREAK_AFTER_HEADER)
    if _add_page_break_before_row(worksheet, PAGE2_PAGE_BREAK_ROW):
        logging.info(
            "Разрыв страницы перед строкой %s (между 59 и 60).",
            PAGE2_PAGE_BREAK_ROW,
        )


def _setup_page3_page_break(worksheet):
    _remove_page_break_before_row(worksheet, PAGE3_PAGE_BREAK_AFTER_HEADER)
    if _add_page_break_before_row(worksheet, PAGE3_PAGE_BREAK_ROW):
        logging.info(
            "Разрыв страницы перед строкой %s (между 118 и 119).",
            PAGE3_PAGE_BREAK_ROW,
        )


def _setup_page_breaks(worksheet, works_count):
    if _uses_page2(works_count):
        _setup_page2_page_break(worksheet)
    if _uses_page3(works_count):
        _setup_page3_page_break(worksheet)


# -------------------------------------------------------------------
# Заполнение строк работ
# -------------------------------------------------------------------
def _fill_one_work_row(ws, row_1based, idx, work):
    r0 = row_1based - 1
    ws.cells.get(r0, COL_NUM).put_value(idx)
    ws.cells.get(r0, COL_NAME).put_value(work.get("name", "   "))
    ws.cells.get(r0, COL_CODE).put_value(work.get("code", "   "))

    cost_start = float(work.get("cost_from_start", 0))
    cost_year = float(work.get("cost_from_year", 0))
    cost_period = float(work.get("cost_report_period", 0))

    set_cell_number_format(ws, r0, COL_COST_START, cost_start, "#,##0.00")
    set_cell_number_format(ws, r0, COL_COST_YEAR, cost_year, "#,##0.00")
    set_cell_number_format(ws, r0, COL_COST_PERIOD, cost_period, "#,##0.00")


def _ensure_signature_block_on_one_page(worksheet, rows_deleted, works_count):
    if rows_deleted <= 0:
        return

    sig_start = SIGNATURE_BLOCK_START_ROW - rows_deleted
    sig_end = SIGNATURE_BLOCK_END_ROW - rows_deleted

    if _signature_can_stay_with_works(works_count):
        return

    if _signature_must_be_on_next_page(works_count):
        if _add_page_break_before_row(worksheet, sig_start):
            logging.info(
                "Блок подписей на следующей странице (%s работ, разрыв перед строкой %s).",
                works_count,
                sig_start,
            )
        return

    manual_breaks = sorted(br.row for br in worksheet.horizontal_page_breaks)
    page_height = _printable_page_height_pt(worksheet)

    sig_height = _rows_height_pt(worksheet, sig_start, sig_end)
    remaining = _remaining_height_before_row(
        worksheet, sig_start, manual_breaks, page_height
    )

    is_split = _row_block_is_split(worksheet, sig_start, sig_end)
    not_enough_space = sig_height > remaining - SIGNATURE_PAGE_SAFETY_MARGIN

    if not is_split and not not_enough_space:
        return

    if _add_page_break_before_row(worksheet, sig_start):
        logging.info(
            "Блок подписей перенесён на следующую страницу (разрыв перед строкой %s).",
            sig_start,
        )


def _row_is_empty(worksheet, row0):
    """Проверяет, что строка не содержит значений (кроме оформления)."""
    max_col = min(worksheet.cells.max_column + 1, 50)
    for col in range(max_col):
        cell = worksheet.cells.get(row0, col)
        if cell.value not in (None, ""):
            return False
    return True


def _remove_empty_row_before_totals(worksheet):
    """Убирает пустую строку между таблицей работ и блоком «Итого»."""
    max_row = worksheet.cells.max_row
    for row0 in range(max_row + 1):
        cell = worksheet.cells.get(row0, LABEL_COL)
        if not cell.value or TOTAL_LABELS["total"] not in str(cell.value):
            continue
        if row0 <= 0:
            return
        prev_row = row0 - 1
        if _row_is_empty(worksheet, prev_row):
            worksheet.cells.delete_rows(prev_row, 1, True)
            logging.info(
                "Удалена пустая строка %s перед итогами.",
                prev_row + 1,
            )
        return


def _compact_worksheet_after_fill(worksheet, works_count, all_rows):
    if works_count >= len(all_rows) or works_count <= 0:
        return 0

    last_work_row = all_rows[works_count - 1]
    last_delete_row = TOTAL_BLOCK_START_ROW - 1
    rows_deleted = 0

    if last_work_row <= PAGE3_LAST_WORK_ROW and _uses_page3(works_count):
        if last_work_row < PAGE3_LAST_WORK_ROW:
            empty_count = PAGE3_LAST_WORK_ROW - last_work_row
            worksheet.cells.delete_rows(last_work_row, empty_count, True)
            rows_deleted += empty_count

        logging.info(
            "Сжатие КС-3 (3-я стр.): удалено %s строк, итоги у строки %s.",
            rows_deleted,
            last_work_row + 1,
        )
        return rows_deleted

    if last_work_row <= PAGE2_LAST_WORK_ROW and _uses_page2(works_count):
        first_delete_row = PAGE3_HEADER_START
        if first_delete_row <= last_delete_row:
            delete_count = last_delete_row - first_delete_row + 1
            worksheet.cells.delete_rows(first_delete_row - 1, delete_count, True)
            rows_deleted += delete_count

        if last_work_row < PAGE2_LAST_WORK_ROW:
            empty_count = PAGE2_LAST_WORK_ROW - last_work_row
            worksheet.cells.delete_rows(last_work_row, empty_count, True)
            rows_deleted += empty_count

        logging.info(
            "Сжатие КС-3 (2-я стр.): удалено %s строк, блоки %s–%s и %s–%s сохранены.",
            rows_deleted,
            PAGE2_HEADER_START,
            PAGE2_HEADER_END,
            PAGE3_HEADER_START,
            PAGE3_HEADER_END,
        )
        return rows_deleted

    first_delete_row = last_work_row + 1
    if first_delete_row > last_delete_row:
        return 0

    delete_count = last_delete_row - first_delete_row + 1
    worksheet.cells.delete_rows(first_delete_row - 1, delete_count, True)
    logging.info(
        "Сжатие КС-3: удалено %s строк (%s–%s), итоги у строки %s.",
        delete_count,
        first_delete_row,
        last_delete_row,
        last_work_row + 1,
    )
    return delete_count


# -------------------------------------------------------------------
# Заполнение шапки, подписей и итогов
# -------------------------------------------------------------------
def _fill_ks3_header_and_signature(worksheet, data):
    """
    Заполняет шапку документа КС-3 и блок подписей.
    Ключи data совпадают с download_act в views.py.
    """
    worksheet.cells.get("F7").put_value(data.get("investor", "   "))
    worksheet.cells.get("AP6").put_value(data.get("okpo_investor", "   "))
    worksheet.cells.get("M9").put_value(data.get("customer", "   "))
    worksheet.cells.get("AP8").put_value(data.get("okpo_customer", "   "))
    worksheet.cells.get("N11").put_value(data.get("contractor", "   "))
    worksheet.cells.get("AP10").put_value(data.get("okpo_contractor", "   "))
    worksheet.cells.get("E13").put_value(data.get("construction", "   "))
    worksheet.cells.get("AP14").put_value(data.get("okdp", "   "))
    worksheet.cells.get("AP16").put_value(data.get("contract_number", "   "))
    worksheet.cells.get("AP17").put_value(data.get("day_contract", "   "))
    worksheet.cells.get("AT17").put_value(data.get("month_contract", "   "))
    worksheet.cells.get("AX17").put_value(data.get("year_contract", "   "))
    worksheet.cells.get("X22").put_value(data.get("document_number", "   "))
    worksheet.cells.get("AG22").put_value(
        data.get("report_date", data.get("report_to", "   "))
    )
    worksheet.cells.get("AR22").put_value(data.get("report_from", "   "))
    worksheet.cells.get("AW22").put_value(data.get("report_to", "   "))

    vat_rate = float(str(data.get("vat_rate", "20%")).replace('%', '')) / 100
    worksheet.cells.get("AS165").put_value(
        f"Сумма НДС {int(round(vat_rate * 100))}%"
    )

    # Заказчик (принял) / подрядчик (сдал)
    worksheet.cells.get("N169").put_value(data.get("accept_position", "   "))
    worksheet.cells.get("AJ170").put_value(data.get("accept_signature", "   "))
    worksheet.cells.get("N174").put_value(data.get("surrender_position", "   "))
    worksheet.cells.get("AJ175").put_value(data.get("surrender_signature", "   "))


def _fill_totals(worksheet, data, works_count):
    vat_rate = float(str(data.get("vat_rate", "20%")).replace('%', '')) / 100
    works_list = data.get("works", [])
    total_sum = sum(
        float(works_list[i].get("cost_report_period", 0))
        for i in range(min(works_count, len(works_list)))
    )
    vat_sum = total_sum * vat_rate
    total_with_vat = total_sum + vat_sum

    max_row = worksheet.cells.max_row
    found = {"total": False, "vat": False, "total_with_vat": False}

    for row0 in range(max_row + 1):
        cell = worksheet.cells.get(row0, LABEL_COL)
        if not cell.value:
            continue
        text = str(cell.value).strip()
        if TOTAL_LABELS["total"] in text and not found["total"]:
            set_cell_number_format(
                worksheet, row0, TOTAL_COL, round(total_sum, 2), "#,##0.00"
            )
            logging.info("✓ Заполнено 'Итого' в строке %d: %.2f", row0 + 1, total_sum)
            found["total"] = True
        elif TOTAL_LABELS["vat"] in text and not found["vat"]:
            worksheet.cells.get(row0, LABEL_COL).put_value(
                f"Сумма НДС {int(round(vat_rate * 100))}%"
            )
            set_cell_number_format(
                worksheet, row0, TOTAL_COL, round(vat_sum, 2), "#,##0.00"
            )
            logging.info("✓ Заполнено 'Сумма НДС' в строке %d: %.2f", row0 + 1, vat_sum)
            found["vat"] = True
        elif TOTAL_LABELS["total_with_vat"] in text and not found["total_with_vat"]:
            set_cell_number_format(
                worksheet, row0, TOTAL_COL, round(total_with_vat, 2), "#,##0.00"
            )
            logging.info(
                "✓ Заполнено 'Всего с учетом НДС' в строке %d: %.2f",
                row0 + 1,
                total_with_vat,
            )
            found["total_with_vat"] = True

        if all(found.values()):
            break

    if not found["total"]:
        logging.warning("Не найдена строка с меткой '%s'", TOTAL_LABELS["total"])
    if not found["vat"]:
        logging.warning("Не найдена строка с меткой '%s'", TOTAL_LABELS["vat"])
    if not found["total_with_vat"]:
        logging.warning(
            "Не найдена строка с меткой '%s'", TOTAL_LABELS["total_with_vat"]
        )


def fill_ks3_worksheet(worksheet, works_list):
    """
    Заполняет лист КС-3: работы по страницам, сжатие, подписи.
    works_list – список словарей с ключами:
        name, code, cost_from_start, cost_from_year, cost_report_period

    Returns:
        int: число удалённых строк при сжатии
    """
    all_rows = _get_all_work_rows()
    works_to_fill = works_list[:MAX_WORK_COUNT]
    if len(works_list) > MAX_WORK_COUNT:
        logging.warning(
            "Получено %s работ, будет заполнено только %s.",
            len(works_list),
            MAX_WORK_COUNT,
        )

    for idx, work in enumerate(works_to_fill, start=1):
        row_1based = all_rows[idx - 1]
        try:
            _fill_one_work_row(worksheet, row_1based, idx, work)
        except Exception as e:
            logging.error(
                "Ошибка при заполнении работы №%s (строка %s): %s",
                idx,
                row_1based,
                e,
            )
            logging.debug(traceback.format_exc())

    rows_deleted = _compact_worksheet_after_fill(
        worksheet, len(works_to_fill), all_rows
    )
    _remove_empty_row_before_totals(worksheet)

    if works_to_fill:
        _ensure_signature_block_on_one_page(
            worksheet, rows_deleted, len(works_to_fill)
        )

    return rows_deleted


def fill_ks3(template_path, output_path, data, format='pdf'):
    """
    Заполнение шаблона КС-3 и экспорт в PDF или XLSX.

    Args:
        template_path: путь к ks3.xlsx
        output_path: путь для сохранения результата
        data: dict с данными для заполнения (шапка, подписи, работы)
        format: 'pdf' или 'xlsx'

    Returns:
        str: путь к сгенерированному файлу
    """
    wb = None
    try:
        wb = ac.Workbook(template_path)
        ws = wb.worksheets[0]

        _fill_ks3_header_and_signature(ws, data)

        works_list = data.get("works", [])
        works_count = len(works_list)

        _setup_page_breaks(ws, works_count)
        fill_ks3_worksheet(ws, works_list)
        _fill_totals(ws, data, min(works_count, MAX_WORK_COUNT))

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        if format == 'xlsx':
            wb.save(output_path, SaveFormat.XLSX)
            logger.info(f"XLSX сохранён: {output_path}")
        else:
            wb.save(output_path, SaveFormat.PDF)
            logger.info(f"PDF сохранён: {output_path}")

        logger.info(f"Заполнено работ: {works_count}")
        return output_path

    except Exception as e:
        logger.error(f"Ошибка генерации КС-3: {e}")
        traceback.print_exc()
        raise
    finally:
        if wb:
            del wb
        gc.collect()
