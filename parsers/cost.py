"""
Разбор файла с себестоимостью товаров для массовой загрузки (/importcost).

Принимает .xlsx или .csv с как минимум двумя колонками:
- код товара (заголовок вроде "Код", "Артикул", "SKU")
- себестоимость за штуку (заголовок вроде "Себестоимость", "Цена", "Стоимость")

Названия колонок ищутся гибко (без учёта регистра, по вхождению ключевых слов),
чтобы не заставлять пользователя подгонять файл под жёсткий формат.
Необязательная колонка "Название" — если её нет, подставится пустое значение.
"""
import pandas as pd

CODE_KEYWORDS = ["код", "артикул", "sku"]
COST_KEYWORDS = ["себестоимост", "цена", "стоимост"]
NAME_KEYWORDS = ["назван", "товар", "name"]


def _find_column(columns, keywords):
    for col in columns:
        low = str(col).lower()
        if any(kw in low for kw in keywords):
            return col
    return None


def parse_cost_file(file_path: str) -> list[tuple[str, str, float]]:
    """Возвращает список (код, название_или_None, себестоимость)."""
    if file_path.lower().endswith(".csv"):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    code_col = _find_column(df.columns, CODE_KEYWORDS)
    cost_col = _find_column(df.columns, COST_KEYWORDS)
    name_col = _find_column(df.columns, NAME_KEYWORDS)

    if code_col is None or cost_col is None:
        raise ValueError(
            "Не нашла нужные колонки. Должна быть колонка с кодом товара "
            "(«Код»/«Артикул»/«SKU») и колонка с ценой («Себестоимость»/«Цена»/«Стоимость»)."
        )

    items = []
    for _, row in df.iterrows():
        code = row[code_col]
        cost = row[cost_col]
        if pd.isna(code) or pd.isna(cost):
            continue
        name = row[name_col] if name_col and pd.notna(row.get(name_col)) else None
        items.append((str(code).strip(), name, float(cost)))

    if not items:
        raise ValueError("В файле не нашлось ни одной валидной строки код+цена.")

    return items
