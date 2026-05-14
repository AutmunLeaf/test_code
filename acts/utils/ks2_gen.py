"""
Генератор актов КС-2.
Использует Aspose.Cells для заполнения шаблона и экспорта в PDF/XLSX.
"""
import os
import aspose.cells as ac
from aspose.cells import SaveFormat
import gc
import traceback
import logging

from ..constants import MAX_WORK_ROWS_KS2

logger = logging.getLogger(__name__)

# --- КС-2: разметка шаблона (1-based строки Excel, как в форме) ---
_P1_WORK_FIRST = 32
_P1_WORK_LAST = 37
_CAPACITY_P1 = _P1_WORK_LAST - _P1_WORK_FIRST + 1  # 6

_P2_WORK_FIRST = 45
# Строк таблицы на 2-й стр. в шаблоне: всего работ MAX − слоты на 1-й стр. (без лишней insert_rows, если шаблон уже размечен).
_P2_TEMPLATE_WORK_SLOTS = MAX_WORK_ROWS_KS2 - _CAPACITY_P1
_P2_TEMPLATE_WORK_LAST = _P2_WORK_FIRST + _P2_TEMPLATE_WORK_SLOTS - 1
# Куда вставлять доп. строки таблицы на 2-й стр., если работ больше, чем слотов в шаблоне.
_P2_INSERT_EXTRA_BEFORE_ROW = _P2_TEMPLATE_WORK_LAST + 1
# 3-я стр.: вставка строк таблицы перед подписями (F169 → строка 169).
_P3_INSERT_BEFORE_ROW = 169
# Пустые строки после таблицы на 3-й стр. до блока подписей (итоги с НДС остаются на 2-й).
_P3_TAIL_ROWS_AFTER_WORKS = 2


def set_cell_number_format(ws, row, col, value, fmt):
    """
    Вставляет число в ячейку и задаёт числовой формат.
    
    Args:
        ws: рабочий лист
        row: индекс строки (начиная с 0)
        col: индекс столбца (начиная с 0)
        value: число (float или int)
        fmt: строка формата (например, "0.###", "#,##0.00")
    """
    cell = ws.cells.get(row, col)
    cell.put_value(value)
    style = cell.get_style()
    style.custom = fmt
    cell.set_style(style)


def _fill_one_work_row(ws, row_1based, idx, work):
    """Заполняет одну строку таблицы работ (индексы строк — как в Excel, с 1)."""
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


def _clear_work_block(ws, row_first_1b, row_last_1b, col_max=32):
    for r in range(row_first_1b, row_last_1b + 1):
        for c in range(col_max):
            ws.cells.get(r - 1, c).put_value("  ")


def fill_ks2(template_path, output_path, data, format='pdf'):
    """
    Заполнение шаблона КС-2 и экспорт в PDF или XLSX.
    
    Args:
        template_path: путь к ks2.xlsx
        output_path: путь для сохранения результата
        data: dict с данными для заполнения
        format: 'pdf' или 'xlsx'
    
    Returns:
        str: путь к сгенерированному файлу
    """
    wb = None
    try:
        # 1. Инициализация Aspose.Cells
        wb = ac.Workbook(template_path)
        ws = wb.worksheets[0]

        vat_rate = float(str(data.get("vat_rate", "20%")).replace('%', '')) / 100

        # 2. Шапка (строки выше зоны вставки на 2-й стр. не сдвигаются)
        ws.cells.get("E7").put_value(data.get("investor", "   "))
        ws.cells.get("G9").put_value(data.get("customer", "   "))
        ws.cells.get("H11").put_value(data.get("contractor", "   "))
        ws.cells.get("E13").put_value(data.get("construction", "   "))
        ws.cells.get("C15").put_value(data.get("object", "   "))
        ws.cells.get("N24").put_value(data.get("document_number", "   "))
        ws.cells.get("Q24").put_value(data.get("contract_date", "   "))
        ws.cells.get("W24").put_value(data.get("report_from", "   "))
        ws.cells.get("AA24").put_value(data.get("report_to", "   "))
        ws.cells.get("AD6").put_value(data.get("okpo_investor", "   "))
        ws.cells.get("AD8").put_value(data.get("okpo_customer", "   "))
        ws.cells.get("AD10").put_value(data.get("okpo_contractor", "   "))
        ws.cells.get("AD16").put_value(data.get("okdp", "   "))
        ws.cells.get("AD18").put_value(data.get("contract_number", "   "))
        ws.cells.get("AD19").put_value(data.get("day_contract", "   "))
        ws.cells.get("AF19").put_value(data.get("month_contract", "   "))
        ws.cells.get("AG19").put_value(data.get("year_contract", "   "))
        ws.cells.get("O27").put_value(data.get("smeta", "   "))

        works = list(data.get("works") or [])
        n = len(works)
        capacity_p1 = _CAPACITY_P1

        # Сценарии: 1 — все на 1-й; 2 — помещаются на 1–2 стр. в размеченную таблицу; 5 — переполнение, хвост на 3-й.
        capacity_p2 = _P2_TEMPLATE_WORK_SLOTS
        if n <= capacity_p1:
            scenario = 1
            works_p1, works_p2, works_p3 = works, [], []
        elif n <= capacity_p1 + capacity_p2:
            scenario = 2
            works_p1 = works[:capacity_p1]
            works_p2 = works[capacity_p1:]
            works_p3 = []
        else:
            scenario = 5
            works_p1 = works[:capacity_p1]
            works_p2 = works[capacity_p1:capacity_p1 + capacity_p2]
            works_p3 = works[capacity_p1 + capacity_p2:]

        d2 = max(0, len(works_p2) - _P2_TEMPLATE_WORK_SLOTS) if works_p2 else 0
        if d2:
            ws.cells.insert_rows(_P2_INSERT_EXTRA_BEFORE_ROW - 1, d2)

        base_shift = d2
        d3 = 0
        p3_first_1b = None
        if works_p3:
            d3 = len(works_p3) + _P3_TAIL_ROWS_AFTER_WORKS
            sig_insert_1b = _P3_INSERT_BEFORE_ROW + base_shift
            ws.cells.insert_rows(sig_insert_1b - 1, d3)
            p3_first_1b = sig_insert_1b

        # Очистка и заполнение 1-й стр.
        _clear_work_block(ws, _P1_WORK_FIRST, _P1_WORK_LAST)
        for i, work in enumerate(works_p1, start=1):
            _fill_one_work_row(ws, _P1_WORK_FIRST + i - 1, i, work)

        d1_footer = 0  # сдвиг подписей после вставок сценария 1

        if not works_p2:
            logger.info(" Сценарий 1: Все работы на 1-й странице")

            ws.cells.insert_rows(39, 4)
            ws.cells.copy_rows(ws.cells, 57, 39, 4)
            ws.cells.delete_rows(45, 16)
            d1_footer = 4

            for r in range(39, _P1_WORK_FIRST - 1, -1):
                cell_val = ws.cells.get(r - 1, 5).value
                if cell_val is None or str(cell_val).strip() in ("", " "):
                    ws.cells.delete_rows(r - 1, 1)

            total_p1 = sum(w.get("quantity", 0) * w.get("price", 0) for w in works_p1)
            vat_p1 = total_p1 * vat_rate
            total_with_vat_p1 = total_p1 + vat_p1

            summary_start_p1 = _P1_WORK_FIRST + len(works_p1)
            set_cell_number_format(ws, summary_start_p1 - 1, 29, round(total_p1, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p1, 29, round(total_p1, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p1 + 1, 29, round(vat_p1, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p1 + 2, 29, round(total_with_vat_p1, 2), "#,##0.00")

            r166 = 166 + d1_footer if 166 >= 40 else 166
            r169 = 169 + d1_footer if 169 >= 40 else 169
            r170 = 170 + d1_footer if 170 >= 40 else 170
            r174 = 174 + d1_footer if 174 >= 40 else 174
            r175 = 175 + d1_footer if 175 >= 40 else 175
        else:
            if scenario == 2:
                logger.info(
                    " Сценарий 2: до %s работ на двух страницах (слотов на 2-й: %s)",
                    capacity_p1 + capacity_p2,
                    capacity_p2,
                )
            else:
                logger.info(
                    " Сценарий 5: >%s работ — %s на 2-й, остальное на 3-й",
                    capacity_p1 + capacity_p2,
                    capacity_p2,
                )

            p2_last_1b = _P2_WORK_FIRST + max(_P2_TEMPLATE_WORK_SLOTS + d2, len(works_p2)) - 1
            _clear_work_block(ws, _P2_WORK_FIRST, p2_last_1b)

            start_num_2 = capacity_p1 + 1
            for j, work in enumerate(works_p2):
                _fill_one_work_row(ws, _P2_WORK_FIRST + j, start_num_2 + j, work)

            if works_p3 and p3_first_1b is not None:
                p3_last_1b = p3_first_1b + len(works_p3) - 1
                _clear_work_block(ws, p3_first_1b, p3_last_1b)
                start_num_3 = capacity_p1 + len(works_p2) + 1
                for j, work in enumerate(works_p3):
                    _fill_one_work_row(ws, p3_first_1b + j, start_num_3 + j, work)

            # Удаляем пустые строки внизу зоны работ на 2-й стр.
            for r in range(p2_last_1b, _P2_WORK_FIRST - 1, -1):
                cell_val = ws.cells.get(r - 1, 5).value
                if cell_val is None or str(cell_val).strip() in ("", " "):
                    ws.cells.delete_rows(r - 1, 1)

            for r in range(38, _P1_WORK_FIRST - 1, -1):
                cell_val = ws.cells.get(r - 1, 5).value
                if cell_val is None or str(cell_val).strip() in ("", " "):
                    ws.cells.delete_rows(r - 1, 1)

            total_p1 = sum(w.get("quantity", 0) * w.get("price", 0) for w in works_p1)
            total_p2 = sum(w.get("quantity", 0) * w.get("price", 0) for w in works_p2)
            total_p3 = sum(w.get("quantity", 0) * w.get("price", 0) for w in works_p3)
            grand_total = total_p1 + total_p2 + total_p3
            grand_vat = grand_total * vat_rate
            grand_total_vat = grand_total + grand_vat

            set_cell_number_format(ws, 37, 29, round(total_p1, 2), "#,##0.00")

            summary_start_p2 = _P2_WORK_FIRST + len(works_p2)
            set_cell_number_format(ws, summary_start_p2 - 2, 29, round(total_p2, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p2 - 1, 29, round(grand_total, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p2, 29, round(grand_vat, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p2 + 1, 29, round(grand_total_vat, 2), "#,##0.00")

            r166 = 166 + base_shift
            r169 = 169 + base_shift + (d3 if works_p3 else 0)
            r170 = 170 + base_shift + (d3 if works_p3 else 0)
            r174 = 174 + base_shift + (d3 if works_p3 else 0)
            r175 = 175 + base_shift + (d3 if works_p3 else 0)

        ws.cells.get(f"S{r166}").put_value(f"Сумма НДС {int(round(vat_rate * 100))}%")
        ws.cells.get(f"F{r169}").put_value(data.get("surrender_position", "   "))
        ws.cells.get(f"N{r170}").put_value(data.get("surrender_signature", "   "))
        ws.cells.get(f"F{r174}").put_value(data.get("accept_position", "   "))
        ws.cells.get(f"N{r175}").put_value(data.get("accept_signature", "   "))

        # 5. Экспорт
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        if format == 'xlsx':
            wb.save(output_path, SaveFormat.XLSX)
            logger.info(f" XLSX сохранён: {output_path}")
        else:
            wb.save(output_path, SaveFormat.PDF)
            logger.info(f" PDF сохранён: {output_path}")
        
        logger.info(f" Заполнено работ: {len(works)}")
        return output_path
        
    except Exception as e:
        logger.error(f" Ошибка генерации КС-2: {e}")
        traceback.print_exc()
        raise
    finally:
        if wb:
            del wb
        gc.collect()