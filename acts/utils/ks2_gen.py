"""
Генератор актов КС-2.
Использует Aspose.Cells для заполнения шаблона и экспорта в PDF/XLSX.

Разметка шаблона (строки Excel, 1-based):
  стр.1 — работы 32–37 (6)
  стр.2 — 45–72 (28)
  стр.3 — 80–108 (29)
  стр.4 — 116–144 (29)
  стр.5 — 152–163 (12)
  итоги в шаблоне — 164–167; незаполненные слоты работ удаляются снизу вверх,
  блок итогов и подписи сдвигаются на (последняя_работа + 1).
"""
import gc
import logging
import os
import traceback

import aspose.cells as ac
from aspose.cells import SaveFormat

from ..constants import MAX_WORK_ROWS_KS2

logger = logging.getLogger(__name__)

# (первая_строка, последняя_строка) включительно
WORK_PAGE_BLOCKS = (
    (32, 37),
    (45, 72),
    (80, 108),
    (116, 144),
    (152, 163),
)

_SUMMARY_TEMPLATE_FIRST = 164

# Подписи в шаблоне: номер строки минус первая строка блока итогов (164)
_SIG_ROW_OFFSETS = {
    "vat_label": 2,  # S166
    "surrender_position": 5,  # F169
    "surrender_signature": 6,  # N170
    "accept_position": 10,  # F174
    "accept_signature": 11,  # N175
}


def _flat_work_rows():
    rows = []
    for a, b in WORK_PAGE_BLOCKS:
        rows.extend(range(a, b + 1))
    return rows


_WORK_SLOT_ROWS = _flat_work_rows()
if len(_WORK_SLOT_ROWS) != MAX_WORK_ROWS_KS2:
    raise ValueError(
        f"WORK_PAGE_BLOCKS: {len(_WORK_SLOT_ROWS)} слотов, "
        f"MAX_WORK_ROWS_KS2={MAX_WORK_ROWS_KS2} — сверьте acts/constants.py"
    )


def set_cell_number_format(ws, row, col, value, fmt):
    cell = ws.cells.get(row, col)
    cell.put_value(value)
    style = cell.get_style()
    style.custom = fmt
    cell.set_style(style)


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


def _clear_work_row(ws, row_1based, col_max=32):
    r0 = row_1based - 1
    for c in range(col_max):
        ws.cells.get(r0, c).put_value("  ")


def _clear_all_work_slots(ws):
    for r in _WORK_SLOT_ROWS:
        _clear_work_row(ws, r)


def _page1_capacity():
    return WORK_PAGE_BLOCKS[0][1] - WORK_PAGE_BLOCKS[0][0] + 1


def fill_ks2(template_path, output_path, data, format="pdf"):
    wb = None
    try:
        wb = ac.Workbook(template_path)
        ws = wb.worksheets[0]

        vat_rate = float(str(data.get("vat_rate", "20%")).replace("%", "")) / 100

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
        cap = len(_WORK_SLOT_ROWS)
        if len(works) > cap:
            logger.warning(
                "КС-2: работ %s, в шаблоне слотов %s — будут заполнены только первые %s",
                len(works),
                cap,
                cap,
            )
            works = works[:cap]
        n = len(works)

        _clear_all_work_slots(ws)

        for i, work in enumerate(works, start=1):
            row_1b = _WORK_SLOT_ROWS[i - 1]
            _fill_one_work_row(ws, row_1b, i, work)

        if n < cap:
            for r in sorted(_WORK_SLOT_ROWS[n:], reverse=True):
                ws.cells.delete_rows(r - 1, 1)

        if n == 0:
            summary_first = _SUMMARY_TEMPLATE_FIRST
        else:
            summary_first = _WORK_SLOT_ROWS[n - 1] + 1

        cap1 = _page1_capacity()
        total_p1 = sum(
            w.get("quantity", 0) * w.get("price", 0)
            for w in works[:cap1]
        )
        total_rest = sum(
            w.get("quantity", 0) * w.get("price", 0)
            for w in works[cap1:]
        )
        grand_total = total_p1 + total_rest
        grand_vat = grand_total * vat_rate
        grand_with_vat = grand_total + grand_vat

        s0 = summary_first
        if n == 0:
            z = 0.0
            set_cell_number_format(ws, s0, 29, z, "#,##0.00")
            set_cell_number_format(ws, s0 + 1, 29, z, "#,##0.00")
            set_cell_number_format(ws, s0 + 2, 29, z, "#,##0.00")
            set_cell_number_format(ws, s0 + 3, 29, z, "#,##0.00")
        elif n <= cap1:
            set_cell_number_format(ws, s0, 29, round(grand_total, 2), "#,##0.00")
            set_cell_number_format(ws, s0 + 1, 29, round(grand_total, 2), "#,##0.00")
            set_cell_number_format(ws, s0 + 2, 29, round(grand_vat, 2), "#,##0.00")
            set_cell_number_format(ws, s0 + 3, 29, round(grand_with_vat, 2), "#,##0.00")
        else:
            set_cell_number_format(ws, s0, 29, round(total_rest, 2), "#,##0.00")
            set_cell_number_format(ws, s0 + 1, 29, round(grand_total, 2), "#,##0.00")
            set_cell_number_format(ws, s0 + 2, 29, round(grand_vat, 2), "#,##0.00")
            set_cell_number_format(ws, s0 + 3, 29, round(grand_with_vat, 2), "#,##0.00")
            set_cell_number_format(ws, 37, 29, round(total_p1, 2), "#,##0.00")

        r_vat_label = s0 + _SIG_ROW_OFFSETS["vat_label"]
        r_sur_pos = s0 + _SIG_ROW_OFFSETS["surrender_position"]
        r_sur_sig = s0 + _SIG_ROW_OFFSETS["surrender_signature"]
        r_acc_pos = s0 + _SIG_ROW_OFFSETS["accept_position"]
        r_acc_sig = s0 + _SIG_ROW_OFFSETS["accept_signature"]

        ws.cells.get(f"S{r_vat_label}").put_value(f"Сумма НДС {int(round(vat_rate * 100))}%")
        ws.cells.get(f"F{r_sur_pos}").put_value(data.get("surrender_position", "   "))
        ws.cells.get(f"N{r_sur_sig}").put_value(data.get("surrender_signature", "   "))
        ws.cells.get(f"F{r_acc_pos}").put_value(data.get("accept_position", "   "))
        ws.cells.get(f"N{r_acc_sig}").put_value(data.get("accept_signature", "   "))

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        if format == "xlsx":
            wb.save(output_path, SaveFormat.XLSX)
            logger.info(" XLSX сохранён: %s", output_path)
        else:
            wb.save(output_path, SaveFormat.PDF)
            logger.info(" PDF сохранён: %s", output_path)

        logger.info(" Заполнено работ: %s", n)
        return output_path

    except Exception as e:
        logger.error(" Ошибка генерации КС-2: %s", e)
        traceback.print_exc()
        raise
    finally:
        if wb:
            del wb
        gc.collect()
