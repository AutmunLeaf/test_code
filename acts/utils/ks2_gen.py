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
from django.conf import settings

logger = logging.getLogger(__name__)


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
        
        # 2. Заполнение шапки (твой оригинальный код)
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
        ws.cells.get("S56").put_value(f"Сумма НДС {int(round(vat_rate * 100))}%")
        
        # 3. Заполнение подписей
        ws.cells.get("F59").put_value(data.get("surrender_position", "   "))
        ws.cells.get("N60").put_value(data.get("surrender_signature", "   "))
        ws.cells.get("F64").put_value(data.get("accept_position", "   "))
        ws.cells.get("N65").put_value(data.get("accept_signature", "   "))
        
        # 4. Разделение и заполнение работ
        start_row_1 = 32
        end_row_1_data = 37
        capacity_p1 = end_row_1_data - start_row_1 + 1  # 6 строк
        
        works = data.get("works", [])
        works_p1 = works[:capacity_p1]
        works_p2 = works[capacity_p1:]
        
        # Очищаем области данных
        for r in range(start_row_1, end_row_1_data + 1):
            for c in range(32):
                ws.cells.get(r - 1, c).put_value("  ")
        for r in range(45, 54):
            for c in range(32):
                ws.cells.get(r - 1, c).put_value("  ")
        
        # Заполняем 1-ю страницу
        row = start_row_1
        for i, work in enumerate(works_p1, start=1):
            ws.cells.get(row - 1, 0).put_value(i)
            ws.cells.get(row - 1, 2).put_value(work.get("position", "   "))
            ws.cells.get(row - 1, 5).put_value(work.get("name", "   "))
            ws.cells.get(row - 1, 14).put_value(work.get("number_pricelist", "   "))
            ws.cells.get(row - 1, 16).put_value(work.get("unit_of_measurement", "   "))
            
            # Количество, цена, сумма – с форматированием
            qty = float(work.get("quantity", 0))
            price = float(work.get("price", 0))
            set_cell_number_format(ws, row - 1, 20, qty, "0.000")
            set_cell_number_format(ws, row - 1, 24, price, "#,##0.00")
            set_cell_number_format(ws, row - 1, 29, qty * price, "#,##0.00")
            
            row += 1
        
        # Логика обработки страниц
        if not works_p2:
            # Сценарий 1: Все работы на 1-й странице
            logger.info(" Сценарий 1: Все работы на 1-й странице")
            
            ws.cells.insert_rows(39, 4)
            ws.cells.copy_rows(ws.cells, 57, 39, 4)
            ws.cells.delete_rows(45, 16)
            
            # Удаляем пустые строки
            for r in range(39, start_row_1 - 1, -1):
                cell_val = ws.cells.get(r - 1, 5).value
                if cell_val is None or str(cell_val).strip() in ("", " "):
                    ws.cells.delete_rows(r - 1, 1)
            
            # Подсчёт итогов
            total_p1 = sum(w.get("quantity", 0) * w.get("price", 0) for w in works_p1)
            vat_p1 = total_p1 * vat_rate
            total_with_vat_p1 = total_p1 + vat_p1
            
            summary_start_p1 = start_row_1 + len(works_p1)
            set_cell_number_format(ws, summary_start_p1 - 1, 29, round(total_p1, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p1, 29, round(total_p1, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p1 + 1, 29, round(vat_p1, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p1 + 2, 29, round(total_with_vat_p1, 2), "#,##0.00")
            
        else:
            # Сценарий 2: Работы на двух страницах
            logger.info(" Сценарий 2: Работы переносятся на 2-ю страницу")
            
            start_row_2 = 45
            row_2 = start_row_2
            start_num_2 = capacity_p1 + 1
            
            for i, work in enumerate(works_p2, start=start_num_2):
                ws.cells.get(row_2 - 1, 0).put_value(i)
                ws.cells.get(row_2 - 1, 2).put_value(work.get("position", "   "))
                ws.cells.get(row_2 - 1, 5).put_value(work.get("name", "   "))
                ws.cells.get(row_2 - 1, 14).put_value(work.get("number_pricelist", "   "))
                ws.cells.get(row_2 - 1, 16).put_value(work.get("unit_of_measurement", "   "))
                
                qty2 = float(work.get("quantity", 0))
                price2 = float(work.get("price", 0))
                set_cell_number_format(ws, row_2 - 1, 20, qty2, "0.000")
                set_cell_number_format(ws, row_2 - 1, 24, price2, "#,##0.00")
                set_cell_number_format(ws, row_2 - 1, 29, qty2 * price2, "#,##0.00")
                
                row_2 += 1
            
            # Удаляем пустые строки на странице 2
            for r in range(53, start_row_2 - 1, -1):
                cell_val = ws.cells.get(r - 1, 5).value
                if cell_val is None or str(cell_val).strip() in ("", " "):
                    ws.cells.delete_rows(r - 1, 1)
            # Удаляем пустые строки на странице 1
            for r in range(38, start_row_1 - 1, -1):
                cell_val = ws.cells.get(r - 1, 5).value
                if cell_val is None or str(cell_val).strip() in ("", " "):
                    ws.cells.delete_rows(r - 1, 1)
            
            # Подсчёт итогов
            total_p1 = sum(w.get("quantity", 0) * w.get("price", 0) for w in works_p1)
            total_p2 = sum(w.get("quantity", 0) * w.get("price", 0) for w in works_p2)

            # Итог по первой странице (ячейка 37,29 – строка 37, столбец 29)
            set_cell_number_format(ws, 36, 29, round(total_p1, 2), "#,##0.00")
            
            grand_total = total_p1 + total_p2
            grand_vat = grand_total * vat_rate
            grand_total_vat = grand_total + grand_vat
            
            summary_start_p2 = start_row_2 + len(works_p2)
            set_cell_number_format(ws, summary_start_p2 - 2, 29, round(total_p2, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p2 - 1, 29, round(grand_total, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p2, 29, round(grand_vat, 2), "#,##0.00")
            set_cell_number_format(ws, summary_start_p2 + 1, 29, round(grand_total_vat, 2), "#,##0.00")
        
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