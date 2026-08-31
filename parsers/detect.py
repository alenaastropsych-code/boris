"""Определяет, что за файл прислали: отчёт Ozon, отчёт ВБ или что-то незнакомое."""
import zipfile
import pandas as pd


def detect_report_type(file_path: str) -> str:
    """Возвращает 'ozon', 'wb' или 'unknown'."""
    lower = file_path.lower()

    if lower.endswith(".zip"):
        try:
            with zipfile.ZipFile(file_path) as z:
                names = [n for n in z.namelist() if n.lower().endswith(".xlsx")]
                if names:
                    return "wb"
        except zipfile.BadZipFile:
            return "unknown"
        return "unknown"

    if lower.endswith(".xlsx"):
        try:
            # Ozon: заголовок "Период: ..." в первой строке, реальные заголовки во второй
            df_head = pd.read_excel(file_path, header=None, nrows=2)
            first_cell = str(df_head.iloc[0, 0])
            if "Начисления" in first_cell or "Период" in first_cell:
                return "ozon"
            # ВБ: первая строка сразу заголовки, среди них "Обоснование для оплаты"
            df_head2 = pd.read_excel(file_path, header=0, nrows=1)
            if "Обоснование для оплаты" in df_head2.columns:
                return "wb"
        except Exception:
            return "unknown"

    return "unknown"
