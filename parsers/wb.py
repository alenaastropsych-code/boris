"""
Парсер "Еженедельного детализированного отчёта" Wildberries (.xlsx внутри .zip).

Логика проверена на реальных отчётах пользователя:
- "К перечислению Продавцу за реализованный Товар" — это уже НЕТТО-выручка
  (после вычета комиссии ВБ и эквайринга), поэтому для строк "Продажа" берём
  только эту колонку и не складываем остальные (иначе задвоение).
- Строки "Возмещение издержек..." и "Возмещение за выдачу и возврат..." почти
  всегда взаимно гасятся (комиссия ВБ + НДС + возмещение ~ 0) — суммируем все
  три компонента как есть.
- "Сумма заказа до комиссии" = Цена розничная с учётом согласованной скидки * Кол-во,
  только для строк "Продажа" — нужна для справочной строки % комиссии ВБ.
"""
import zipfile
import tempfile
import os
import hashlib
import json
import pandas as pd

CATEGORY_MAP = {
    "Продажа": "Продажи",
    "Возврат": "Возврат",
    "Логистика": "Логистика",
    "Обработка товара": "Обработка товара",
    "Хранение": "Хранение",
    "Коррекция хранения": "Хранение",
    "Удержание": "Удержания",
    "Штраф": "Штрафы",
    "Возмещение издержек по перевозке/по складским операциям с товаром": "Возмещения логистики",
    "Возмещение за выдачу и возврат товаров на ПВЗ": "Возмещения ПВЗ",
    "Компенсация скидки по программе лояльности": "Программа лояльности",
}


def _extract_xlsx(file_path: str) -> str:
    """Если пришёл zip — достаёт xlsx во временную папку и возвращает путь к нему."""
    if file_path.lower().endswith(".zip"):
        tmpdir = tempfile.mkdtemp()
        with zipfile.ZipFile(file_path) as z:
            names = [n for n in z.namelist() if n.lower().endswith(".xlsx")]
            if not names:
                raise ValueError("В zip-архиве не найден xlsx-файл отчёта ВБ")
            z.extract(names[0], tmpdir)
            return os.path.join(tmpdir, names[0])
    return file_path


def _row_amount(r) -> float:
    ob = r["Обоснование для оплаты"]
    if ob == "Продажа":
        return float(r["К перечислению Продавцу за реализованный Товар"] or 0)
    if ob == "Возврат":
        # Строка с "Тип документа" = Возврат — это отмена/возврат ранее оплаченного
        # заказа: сумма в той же колонке, что и у продажи, но её нужно вычесть,
        # а не прибавить (иначе возврат товара увеличивал бы доход).
        return -float(r["К перечислению Продавцу за реализованный Товар"] or 0)
    if ob == "Логистика":
        return -float(r["Услуги по доставке товара покупателю"] or 0)
    if ob == "Обработка товара":
        return -float(r["Операции на приемке"] or 0)
    if ob in ("Хранение", "Коррекция хранения"):
        return -float(r["Хранение"] or 0)
    if ob == "Удержание":
        # Знак в колонке "Удержания" в отчёте ВБ инвертирован относительно смысла:
        # положительное значение = списание с продавца (расход), отрицательное =
        # выплата продавцу, например компенсация (доход). Поэтому знак переворачиваем.
        return -float(r["Удержания"] or 0)
    if ob == "Штраф":
        # "Общая сумма штрафов" — положительное число означает сумму штрафа,
        # удержанную с продавца, то есть расход. Вычитаем.
        return -float(r.get("Общая сумма штрафов", 0) or 0)
    if ob == "Возмещение издержек по перевозке/по складским операциям с товаром":
        return (float(r["Возмещение издержек по перевозке/по складским операциям с товаром"] or 0)
                + float(r["Вознаграждение Вайлдберриз (ВВ), без НДС"] or 0)
                + float(r["НДС с Вознаграждения Вайлдберриз"] or 0))
    if ob == "Возмещение за выдачу и возврат товаров на ПВЗ":
        return (float(r["Возмещение за выдачу и возврат товаров на ПВЗ"] or 0)
                + float(r["Вознаграждение Вайлдберриз (ВВ), без НДС"] or 0)
                + float(r["НДС с Вознаграждения Вайлдберриз"] or 0))
    if ob == "Компенсация скидки по программе лояльности":
        return float(r["Компенсация скидки по программе лояльности"] or 0)
    # неизвестный тип операции — берём "Общая сумма штрафов", если она есть,
    # но считаем это расходом по умолчанию (безопаснее, чем случайно завысить доход)
    return -float(r.get("Общая сумма штрафов", 0) or 0)


def parse_wb_report(file_path: str, source_name: str) -> list[dict]:
    xlsx_path = _extract_xlsx(file_path)
    df = pd.read_excel(xlsx_path, sheet_name="Sheet1", header=0)

    rows = []
    for _, r in df.iterrows():
        ob = r["Обоснование для оплаты"]
        if pd.isna(ob):
            continue
        group = CATEGORY_MAP.get(ob, ob)
        tx_date = pd.to_datetime(r["Дата продажи"]).date()
        amount = _row_amount(r)

        article = None
        qty = 0.0
        gross = 0.0
        if ob == "Продажа":
            article = str(r["Код номенклатуры"]) if pd.notna(r["Код номенклатуры"]) else None
            qty = float(r["Кол-во"]) if pd.notna(r["Кол-во"]) else 0.0
            price = float(r["Цена розничная с учетом согласованной скидки"] or 0)
            gross = price * qty

        rows.append({
            "platform": "wb",
            "tx_date": tx_date.isoformat(),
            "group_name": group,
            "article": article,
            "qty": qty,
            "amount": amount,
            "gross_amount": gross,
            "source_file": source_name,
        })
    return rows


def get_period(file_path: str):
    xlsx_path = _extract_xlsx(file_path)
    df = pd.read_excel(xlsx_path, sheet_name="Sheet1", header=0)
    dates = pd.to_datetime(df["Дата продажи"]).dropna()
    return dates.min().date(), dates.max().date()


def content_hash(rows: list[dict]) -> str:
    """Хэш данных отчёта (без имени файла) — чтобы ловить повторную загрузку
    того же отчёта под другим именем файла (частый случай при повторном
    скачивании с площадки)."""
    normalized = [{k: v for k, v in r.items() if k != "source_file"} for r in rows]
    normalized.sort(key=lambda r: json.dumps(r, sort_keys=True, default=str))
    blob = json.dumps(normalized, sort_keys=True, default=str)
    return hashlib.sha256(blob.encode()).hexdigest()
