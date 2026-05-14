"""Лимиты строк работ: форма ввода и генерация печатных форм."""

# КС-2: шаблон и ks2_gen — до MAX_WORK_ROWS_KS2 строк (6 на 1-й стр., остальное на 2-й и при переполнении на 3-й).
MAX_WORK_ROWS_KS2 = 100

# КС-3: в шаблоне 14 слотов под строки; fill_ks3 при большем числе вставляет строки
# перед блоком итогов и сдвигает НДС/подписи. Очень большие значения — риск для PDF/памяти.
MAX_WORK_ROWS_KS3 = 100


def max_work_rows_for_act_type(act_type: str) -> int:
    if act_type == "ks3":
        return MAX_WORK_ROWS_KS3
    return MAX_WORK_ROWS_KS2
