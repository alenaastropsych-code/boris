"""
Парсер отчёта Ozon "Отчёт по начислениям" (.xlsx).

Логика проверена на реальном отчёте пользователя:
- Строки группируются по колонке "Группа услуг" — это и есть статьи PnL.
- Себестоимость считаем только по строкам Тип начисления == "Выручка"
  (внутри группы "Продажи" один заказ даёт 3 строки: Выручка, Программы
  партнёров, Баллы за скидки — с одинаковым количеством, поэтому для
  количества товара используем только "Выручка", иначе будет утроение).
"""
import pandas as pd


def parse_ozon_report(file_path: str, source_name: str) -> list[dict]:
    df = pd.read_excel(file_path, header=1)

    rows = []
    for _, r in df.iterrows():
        group = r["Группа услуг"]
        if pd.isna(group):
            continue
        tx_date = pd.to_datetime(r["Дата начисления"]).date()
        amount = float(r["Сумма итого, руб."]) if pd.notna(r["Сумма итого, руб."]) else 0.0

        article = None
        qty = 0.0
        if group == "Продажи" and r.get("Тип начисления") == "Выручка":
            article = str(r["Артикул"]) if pd.notna(r["Артикул"]) else None
            qty = float(r["Количество"]) if pd.notna(r["Количество"]) else 0.0

        rows.append({
            "platform": "ozon",
            "tx_date": tx_date.isoformat(),
            "group_name": group,
            "article": article,
            "qty": qty,
            "amount": amount,
            "gross_amount": 0.0,
            "source_file": source_name,
        })
    return rows


def get_period(file_path: str):
    df = pd.read_excel(file_path, header=1)
    dates = pd.to_datetime(df["Дата начисления"]).dropna()
    return dates.min().date(), dates.max().date()
