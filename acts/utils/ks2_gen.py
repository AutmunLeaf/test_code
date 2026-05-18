"""
Генератор актов КС-2.
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
    number_format – строка формата, например "0.000"
    """
    cell = ws.cells.get(row0, col0)
    cell.put_value(value)
    style = cell.get_style()
    style.custom = number_format
    cell.set_style(style)


# -------------------------------------------------------------------
# Заполнение одной строки с работой (дано в условии)
# -------------------------------------------------------------------
def _fill_one_work_row(ws, row_1based, idx, work):
    r0 = row_1based - 1
    ws.cells.get(r0, 0).put_value(idx)
    ws.cells.get(r0, 2).put_value(work.get("position", "   "))
    ws.cells.get(r0, 5).put_value(work.get("name", "   "))
    ws.cells.get(r0, 14).put_value(work.get("number_pricelist", "   "))
    ws.cells.get(r0, 16).put_value(work.get("unit_of_measurement", "   "))
    qty = float(work.get("quantity", 0))
    price = float(work.get("price", 0))
    set_cell_number_format(ws, r0, 20, qty, "0.000")
    set_cell_number_format(ws, r0, 24, price, "#,##0.00")
    set_cell_number_format(ws, r0, 29, qty * price, "#,##0.00")


# Строки блока итогов в шаблоне (1-based): Итого, Всего по акту, Сумма НДС, Всего с учётом НДС
TOTAL_BLOCK_START_ROW = 164
TOTAL_BLOCK_ROW_COUNT = 4
# Блок подписей в шаблоне (1-based)
SIGNATURE_BLOCK_START_ROW = 169
SIGNATURE_BLOCK_END_ROW = 177
# Границы страниц формы (1-based)
PAGE1_LAST_WORK_ROW = 37
PAGE2_FIRST_WORK_ROW = 45
PAGE2_LAST_WORK_ROW = 72
PAGE3_FIRST_WORK_ROW = 80
PAGE3_LAST_WORK_ROW = 108
PAGE4_FIRST_WORK_ROW = 116
PAGE4_LAST_WORK_ROW = 144
PAGE5_FIRST_WORK_ROW = 152
PAGE5_LAST_WORK_ROW = 163
# 32–34 работы — итоги на 3-й странице; 61–63 — на 4-й; 91–92 — на 5-й (как при 93+)
WORKS_PAGE3_TOTALS_LAYOUT_MIN = 32
WORKS_PAGE3_TOTALS_LAYOUT_MAX = 34
WORKS_PAGE4_TOTALS_LAYOUT_MIN = 61
WORKS_PAGE4_TOTALS_LAYOUT_MAX = 63
WORKS_PAGE5_TOTALS_LAYOUT_MIN = 91
WORKS_PAGE5_TOTALS_LAYOUT_MAX = 92
PAGE1_FULL_WORKS_COUNT = 6  # все слоты 1-й страницы заполнены
# Запас по высоте страницы (pt), чтобы блок не оказался на границе листа
SIGNATURE_PAGE_SAFETY_MARGIN = 25


# -------------------------------------------------------------------
# Удаление пустых строк и сдвиг блока итогов к последней работе
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
    paper_height = 842.0  # A4, портрет
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
    """Номер страницы внутри текущей секции (между ручными разрывами)."""
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
    """Свободная высота на текущей странице перед указанной строкой."""
    row0 = row_1based - 1
    section_start = _manual_section_start(row0, manual_breaks)

    used_height = 0.0
    for r in range(section_start, row0):
        height = _row_height_pt(worksheet, r)
        if used_height > 0 and used_height + height > page_height:
            used_height = 0.0
        used_height += height

    return page_height - used_height


def _row_block_is_split(worksheet, block_start, block_end):
    manual_breaks = sorted(br.row for br in worksheet.horizontal_page_breaks)
    page_height = _printable_page_height_pt(worksheet)
    pages = {
        _print_page_in_section(worksheet, row, page_height, manual_breaks)
        for row in range(block_start, block_end + 1)
    }
    return len(pages) > 1


def _add_page_break_before_row(worksheet, row_1based):
    """Разрыв страницы перед строкой row_1based."""
    break_row = row_1based - 1
    for existing in worksheet.horizontal_page_breaks:
        if existing.row == break_row:
            return False

    worksheet.horizontal_page_breaks.add(break_row)
    return True


def _remove_page_break_before_row(worksheet, row_1based):
    """Убирает разрыв страницы перед строкой, если он есть."""
    break_row = row_1based - 1
    breaks = worksheet.horizontal_page_breaks
    to_remove = [i for i, br in enumerate(breaks) if br.row == break_row]
    for i in reversed(to_remove):
        breaks.remove_at(i)
    return bool(to_remove)


def _uses_page2_totals_layout(works_count, last_work_row):
    """6 работ: блок итогов на 2-й странице формы (после шапки стр. 2)."""
    return (
        works_count == PAGE1_FULL_WORKS_COUNT
        and last_work_row <= PAGE1_LAST_WORK_ROW
    )


def _uses_page3_totals_layout(works_count, last_work_row):
    """32–34 работы: блок итогов на 3-й странице формы (после шапки стр. 3)."""
    return (
        WORKS_PAGE3_TOTALS_LAYOUT_MIN
        <= works_count
        <= WORKS_PAGE3_TOTALS_LAYOUT_MAX
        and last_work_row <= PAGE2_LAST_WORK_ROW
    )


def _uses_page4_totals_layout(works_count, last_work_row):
    """61–63 работы: блок итогов на 4-й странице формы (после шапки стр. 4)."""
    return (
        WORKS_PAGE4_TOTALS_LAYOUT_MIN
        <= works_count
        <= WORKS_PAGE4_TOTALS_LAYOUT_MAX
        and last_work_row <= PAGE3_LAST_WORK_ROW
    )


def _uses_page5_totals_layout(works_count, last_work_row):
    """91–92 работы: блок итогов на 5-й странице формы (после шапки стр. 5)."""
    return (
        WORKS_PAGE5_TOTALS_LAYOUT_MIN
        <= works_count
        <= WORKS_PAGE5_TOTALS_LAYOUT_MAX
        and last_work_row <= PAGE4_LAST_WORK_ROW
    )


def _uses_relocated_totals_layout(works_count, last_work_row):
    return (
        _uses_page2_totals_layout(works_count, last_work_row)
        or _uses_page3_totals_layout(works_count, last_work_row)
        or _uses_page4_totals_layout(works_count, last_work_row)
        or _uses_page5_totals_layout(works_count, last_work_row)
    )


def _should_delete_trailing_page2_work_rows(last_work_row):
    """Пустые слоты 2-й страницы — только если работы там не заполнялись до конца."""
    return PAGE2_FIRST_WORK_ROW <= last_work_row < PAGE2_LAST_WORK_ROW


def _should_delete_trailing_page3_work_rows(last_work_row):
    """Пустые слоты 3-й страницы — только если работы там не заполнялись до конца."""
    return PAGE3_FIRST_WORK_ROW <= last_work_row < PAGE3_LAST_WORK_ROW


def _should_delete_trailing_page4_work_rows(last_work_row):
    """Пустые слоты 4-й страницы — только если работы там не заполнялись до конца."""
    return PAGE4_FIRST_WORK_ROW <= last_work_row < PAGE4_LAST_WORK_ROW


def _ensure_totals_stay_with_works(
    worksheet, last_work_row, total_start, total_end
):
    """
    Четыре строки итогов остаются на одной странице с последней работой
    и не разрываются между страницами.
    """
    manual_breaks = sorted(br.row for br in worksheet.horizontal_page_breaks)
    page_height = _printable_page_height_pt(worksheet)

    # Убираем ошибочный разрыв между работами и итогами
    for row in range(last_work_row + 1, total_end + 1):
        _remove_page_break_before_row(worksheet, row)

    last_work_page = _print_page_in_section(
        worksheet, last_work_row, page_height, manual_breaks
    )
    total_start_page = _print_page_in_section(
        worksheet, total_start, page_height, manual_breaks
    )
    totals_split = _row_block_is_split(worksheet, total_start, total_end)

    if last_work_page == total_start_page and not totals_split:
        return

    group_height = _rows_height_pt(worksheet, last_work_row, total_end)
    if group_height > page_height:
        logging.warning(
            "Последняя работа и итоги не помещаются на одну страницу (%.0f pt > %.0f pt).",
            group_height,
            page_height,
        )
        return

    if _add_page_break_before_row(worksheet, last_work_row):
        logging.info(
            "Итоги закреплены за последней работой (разрыв перед строкой %s).",
            last_work_row,
        )


def _ensure_signature_block_on_one_page(
    worksheet, rows_deleted, last_work_row, works_count
):
    if rows_deleted <= 0:
        return

    total_start = TOTAL_BLOCK_START_ROW - rows_deleted
    total_end = total_start + TOTAL_BLOCK_ROW_COUNT - 1
    sig_start = SIGNATURE_BLOCK_START_ROW - rows_deleted
    sig_end = SIGNATURE_BLOCK_END_ROW - rows_deleted

    if not _uses_relocated_totals_layout(works_count, last_work_row):
        _ensure_totals_stay_with_works(
            worksheet, last_work_row, total_start, total_end
        )

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


def _compact_worksheet_after_fill(worksheet, works_count, all_rows):
    """
    Если работ меньше 104, удаляет пустые строки до блока итогов.
    6 раб. — итоги на 2-й стр.; 32–34 — на 3-й; 61–63 — на 4-й; 91–92 — на 5-й.

    Возвращает число удалённых строк (0, если сжатие не выполнялось).
    """
    if works_count >= len(all_rows) or works_count <= 0:
        return 0

    last_work_row = all_rows[works_count - 1]
    last_delete_row = TOTAL_BLOCK_START_ROW - 1

    if _uses_page2_totals_layout(works_count, last_work_row):
        first_delete_row = PAGE2_FIRST_WORK_ROW
        if first_delete_row > last_delete_row:
            return 0

        delete_count = last_delete_row - first_delete_row + 1
        worksheet.cells.delete_rows(
            first_delete_row - 1,
            delete_count,
            True,
        )
        total_row = TOTAL_BLOCK_START_ROW - delete_count
        logging.info(
            "Сжатие (6 раб.): удалено %s строк, итоги на 2-й странице (строка %s).",
            delete_count,
            total_row,
        )
        return delete_count

    if _uses_page3_totals_layout(works_count, last_work_row):
        first_delete_row = PAGE3_FIRST_WORK_ROW
        if first_delete_row > last_delete_row:
            return 0

        delete_count = last_delete_row - first_delete_row + 1
        worksheet.cells.delete_rows(
            first_delete_row - 1,
            delete_count,
            True,
        )
        rows_deleted = delete_count

        if _should_delete_trailing_page2_work_rows(last_work_row):
            empty_count = PAGE2_LAST_WORK_ROW - last_work_row
            worksheet.cells.delete_rows(
                last_work_row,
                empty_count,
                True,
            )
            rows_deleted += empty_count

        total_row = TOTAL_BLOCK_START_ROW - rows_deleted
        logging.info(
            "Сжатие (32–34 раб.): удалено %s строк, итоги на 3-й странице (строка %s).",
            rows_deleted,
            total_row,
        )
        return rows_deleted

    if _uses_page4_totals_layout(works_count, last_work_row):
        first_delete_row = PAGE4_FIRST_WORK_ROW
        if first_delete_row > last_delete_row:
            return 0

        delete_count = last_delete_row - first_delete_row + 1
        worksheet.cells.delete_rows(
            first_delete_row - 1,
            delete_count,
            True,
        )
        rows_deleted = delete_count

        if _should_delete_trailing_page3_work_rows(last_work_row):
            empty_count = PAGE3_LAST_WORK_ROW - last_work_row
            worksheet.cells.delete_rows(
                last_work_row,
                empty_count,
                True,
            )
            rows_deleted += empty_count

        total_row = TOTAL_BLOCK_START_ROW - rows_deleted
        logging.info(
            "Сжатие (61–63 раб.): удалено %s строк, итоги на 4-й странице (строка %s).",
            rows_deleted,
            total_row,
        )
        return rows_deleted

    if _uses_page5_totals_layout(works_count, last_work_row):
        first_delete_row = PAGE5_FIRST_WORK_ROW
        if first_delete_row > last_delete_row:
            return 0

        delete_count = last_delete_row - first_delete_row + 1
        worksheet.cells.delete_rows(
            first_delete_row - 1,
            delete_count,
            True,
        )
        rows_deleted = delete_count

        if _should_delete_trailing_page4_work_rows(last_work_row):
            empty_count = PAGE4_LAST_WORK_ROW - last_work_row
            worksheet.cells.delete_rows(
                last_work_row,
                empty_count,
                True,
            )
            rows_deleted += empty_count

        total_row = TOTAL_BLOCK_START_ROW - rows_deleted
        logging.info(
            "Сжатие (91–92 раб.): удалено %s строк, итоги на 5-й странице (строка %s).",
            rows_deleted,
            total_row,
        )
        return rows_deleted

    first_delete_row = last_work_row + 1
    if first_delete_row > last_delete_row:
        return 0

    rows_to_delete = last_delete_row - first_delete_row + 1
    worksheet.cells.delete_rows(
        first_delete_row - 1,
        rows_to_delete,
        True,
    )
    logging.info(
        "Сжатие листа: удалено %s строк (%s–%s), итоги сдвинуты к строке %s.",
        rows_to_delete,
        first_delete_row,
        last_delete_row,
        last_work_row + 1,
    )
    return rows_to_delete


# -------------------------------------------------------------------
# Основная функция заполнения листа
# -------------------------------------------------------------------
def _fill_ks2_header_and_signature(worksheet, data):
    """
    Заполняет шапку документа КС-2 и блок подписей.
    
    Args:
        worksheet: объект Worksheet
        data: dict с данными для заполнения
    """
    vat_rate = float(str(data.get("vat_rate", "20%")).replace('%', '')) / 100
    
    # Заполнение шапки
    worksheet.cells.get("E7").put_value(data.get("investor", "   "))
    worksheet.cells.get("G9").put_value(data.get("customer", "   "))
    worksheet.cells.get("H11").put_value(data.get("contractor", "   "))
    worksheet.cells.get("E13").put_value(data.get("construction", "   "))
    worksheet.cells.get("C15").put_value(data.get("object", "   "))
    worksheet.cells.get("N24").put_value(data.get("document_number", "   "))
    worksheet.cells.get("Q24").put_value(data.get("contract_date", "   "))
    worksheet.cells.get("W24").put_value(data.get("report_from", "   "))
    worksheet.cells.get("AA24").put_value(data.get("report_to", "   "))
    worksheet.cells.get("AD6").put_value(data.get("okpo_investor", "   "))
    worksheet.cells.get("AD8").put_value(data.get("okpo_customer", "   "))
    worksheet.cells.get("AD10").put_value(data.get("okpo_contractor", "   "))
    worksheet.cells.get("AD16").put_value(data.get("okdp", "   "))
    worksheet.cells.get("AD18").put_value(data.get("contract_number", "   "))
    worksheet.cells.get("AD19").put_value(data.get("day_contract", "   "))
    worksheet.cells.get("AF19").put_value(data.get("month_contract", "   "))
    worksheet.cells.get("AG19").put_value(data.get("year_contract", "   "))
    worksheet.cells.get("O27").put_value(data.get("smeta", "   "))
    worksheet.cells.get("S166").put_value(f"Сумма НДС {int(round(vat_rate * 100))}%")
    
    # Заполнение подписей
    worksheet.cells.get("F169").put_value(data.get("surrender_position", "   "))
    worksheet.cells.get("N170").put_value(data.get("surrender_signature", "   "))
    worksheet.cells.get("F174").put_value(data.get("accept_position", "   "))
    worksheet.cells.get("N175").put_value(data.get("accept_signature", "   "))


def fill_ks2_worksheet(worksheet, works_list):
    """
    Заполняет лист KS-2 данными из works_list и вычисляет промежуточные итоги для каждой страницы.
    worksheet – объект Worksheet (уже загруженный шаблон)
    works_list – список словарей с ключами:
                 position, name, number_pricelist, unit_of_measurement,
                 quantity, price
    
    Returns:
        tuple: (rows_deleted, page_sums) где page_sums - dict с суммами для каждой страницы
    """
    # Определяем страницы и строки (1-based номера строк шаблона)
    # Всего 104 места
    pages = [
        {"page": 1, "rows": list(range(32, 38)), "total_row": 38},   # 6 работ: строки 32–37, итого строка 38
        {"page": 2, "rows": list(range(45, 73)), "total_row": 73},   # 28 работ: строки 45–72, итого строка 73
        {"page": 3, "rows": list(range(80, 109)), "total_row": 109}, # 29 работ: строки 80–108, итого строка 109
        {"page": 4, "rows": list(range(116, 145)), "total_row": 145}, # 29 работ: строки 116–144, итого строка 145
        {"page": 5, "rows": list(range(152, 164)), "total_row": 164}  # 12 работ: строки 152–163, итого строка 164
    ]

    # Собираем все доступные строки по порядку
    all_rows = []
    for page_info in pages:
        all_rows.extend(page_info["rows"])

    total_rows = len(all_rows)  # 104

    # Ограничиваем количество работ доступным числом строк
    works_to_fill = works_list[:total_rows]
    if len(works_list) > total_rows:
        logging.warning(
            f"Получено {len(works_list)} работ, будет заполнено только {total_rows}."
        )

    # Заполняем построчно
    for idx, work in enumerate(works_to_fill, start=1):
        row_1based = all_rows[idx - 1]
        try:
            _fill_one_work_row(worksheet, row_1based, idx, work)
        except Exception as e:
            logging.error(
                f"Ошибка при заполнении работы №{idx} (строка {row_1based}): {e}"
            )
            logging.debug(traceback.format_exc())

    # Вычисляем промежуточные суммы для каждой страницы ПЕРЕД сжатием
    page_sums = {}
    work_idx = 0
    for page_info in pages:
        page_rows = page_info["rows"]
        page_num = page_info["page"]
        
        # Вычисляем сумму работ для этой страницы
        page_sum = 0.0
        page_start_work_idx = work_idx
        
        for _ in page_rows:
            if work_idx < len(works_to_fill):
                work = works_to_fill[work_idx]
                qty = float(work.get("quantity", 0))
                price = float(work.get("price", 0))
                page_sum += qty * price
                work_idx += 1
            else:
                break
        
        page_sums[page_num] = {
            "sum": round(page_sum, 2),
            "total_row": page_info["total_row"],
            "has_works": work_idx > page_start_work_idx
        }

    rows_deleted = _compact_worksheet_after_fill(
        worksheet, len(works_to_fill), all_rows
    )
    
    if works_to_fill:
        last_work_row = all_rows[len(works_to_fill) - 1]
        _ensure_signature_block_on_one_page(
            worksheet, rows_deleted, last_work_row, len(works_to_fill)
        )
    
    return rows_deleted, page_sums


def fill_ks2(template_path, output_path, data, format='pdf'):
    """
    Заполнение шаблона КС-2 и экспорт в PDF или XLSX.
    
    Args:
        template_path: путь к ks2.xlsx
        output_path: путь для сохранения результата
        data: dict с данными для заполнения (шапка, подписи, работы)
        format: 'pdf' или 'xlsx'
    
    Returns:
        str: путь к сгенерированному файлу
    """
    wb = None
    try:
        # Инициализация Aspose.Cells
        wb = ac.Workbook(template_path)
        ws = wb.worksheets[0]
        
        # Заполнение шапки и подписей
        _fill_ks2_header_and_signature(ws, data)
        
        # Заполнение работ (возвращает количество удалённых строк и суммы по страницам)
        works_list = data.get("works", [])
        rows_deleted, page_sums = fill_ks2_worksheet(ws, works_list)
        
        # Заполняем промежуточные итоги для каждой страницы
        # Ищем строки с текстом "Итого" и заполняем соответствующую ячейку в столбце AD (29)
        # Ограничиваем поиск только строками до строки итогов (164 или меньше)
        max_search_row = 160  # Достаточно для поиска итогов на всех страницах
        
        for page_num, page_info in page_sums.items():
            page_sum = page_info["sum"]
            
            if not page_info["has_works"]:
                continue
            
            # Для каждой страницы ищем строку "Итого"
            # Диапазон поиска зависит от номера страницы
            if page_num == 1:
                search_start, search_end = 30, 40
            elif page_num == 2:
                search_start, search_end = 40, 80
            elif page_num == 3:
                search_start, search_end = 75, 115
            elif page_num == 4:
                search_start, search_end = 110, 150
            else:  # page_num == 5
                search_start, search_end = 145, 165
            
            # Ищем строку "Итого" в столбце S (18) в указанном диапазоне
            found = False
            for row_idx in range(search_start, min(search_end, max_search_row)):
                cell = ws.cells.get(row_idx, 18)  # столбец S (0-based)
                if cell and cell.value and "Итого" in str(cell.value):
                    # Найдена строка Итого, заполняем столбец AD (29) для этой строки
                    set_cell_number_format(
                        ws,
                        row_idx,
                        29,  # столбец AD
                        page_sum,
                        "#,##0.00"
                    )
                    found = True
                    logging.info(f"Страница {page_num}: заполнено Итого в строке {row_idx+1}, сумма={page_sum}")
                    break
            
            if not found:
                logging.warning(f"Страница {page_num}: не найдена строка Итого для заполнения")
        
        # Расчёт и заполнение финальных итогов (Всего, НДС, Всего с НДС)
        vat_rate = float(str(data.get("vat_rate", "20%")).replace('%', '')) / 100
        total_sum = sum(float(w.get("quantity", 0)) * float(w.get("price", 0)) for w in works_list)
        vat_sum = total_sum * vat_rate
        total_with_vat = total_sum + vat_sum
        
        # Ищем строки финальных итогов (Всего по акту, НДС, Всего с НДС)
        # Они находятся в конце документа в столбце S
        final_totals_found = {}
        
        # Ищем с конца документа в обратном порядке для надежности
        max_row = ws.cells.max_row
        search_start = max_row - 1
        search_end = max(0, max_row - 100)
        
        for row_idx in range(search_start, search_end, -1):
            cell = ws.cells.get(row_idx, 18)  # столбец S (0-based)
            if not cell or not cell.value:
                continue
            
            cell_value = str(cell.value).strip()
            
            if "Всего по акту" in cell_value and "total" not in final_totals_found:
                final_totals_found["total"] = row_idx
                logging.info(f"Найдено 'Всего по акту' в строке {row_idx+1}")
            elif "Сумма НДС" in cell_value and "vat" not in final_totals_found:
                final_totals_found["vat"] = row_idx
                logging.info(f"Найдено 'Сумма НДС' в строке {row_idx+1}")
            elif "Всего с учётом НДС" in cell_value and "total_with_vat" not in final_totals_found:
                final_totals_found["total_with_vat"] = row_idx
                logging.info(f"Найдено 'Всего с учётом НДС' в строке {row_idx+1}")
            
            # Если нашли все три, можно остановиться
            if len(final_totals_found) == 3:
                break
        
        # Заполняем финальные итоги
        if "total" in final_totals_found:
            set_cell_number_format(ws, final_totals_found["total"], 29, round(total_sum, 2), "#,##0.00")
            logging.info(f"✓ Заполнено 'Всего по акту': {round(total_sum, 2)}")
        else:
            logging.warning("✗ 'Всего по акту' не найдено!")
        
        if "vat" in final_totals_found:
            set_cell_number_format(ws, final_totals_found["vat"], 29, round(vat_sum, 2), "#,##0.00")
            logging.info(f"✓ Заполнено 'Сумма НДС': {round(vat_sum, 2)}")
        else:
            logging.warning("✗ 'Сумма НДС' не найдено!")
        
        if "total_with_vat" in final_totals_found:
            set_cell_number_format(ws, final_totals_found["total_with_vat"], 29, round(total_with_vat, 2), "#,##0.00")
            logging.info(f"✓ Заполнено 'Всего с учётом НДС': {round(total_with_vat, 2)}")
        else:
            logging.warning("✗ 'Всего с учётом НДС' не найдено!")
        
        # Экспорт
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        if format == 'xlsx':
            wb.save(output_path, SaveFormat.XLSX)
            logger.info(f"XLSX сохранён: {output_path}")
        else:
            wb.save(output_path, SaveFormat.PDF)
            logger.info(f"PDF сохранён: {output_path}")
        
        logger.info(f"Заполнено работ: {len(works_list)}")
        return output_path
        
    except Exception as e:
        logger.error(f"Ошибка генерации КС-2: {e}")
        traceback.print_exc()
        raise
    finally:
        if wb:
            del wb
        gc.collect()
