"""
База данных Бориса — хранит все загруженные операции по Ozon и Wildberries,
себестоимость товаров и лог обработанных файлов.

Используется SQLite (файл data/boris.db) — этого достаточно для объёмов
одного продавца (десятки тысяч строк в месяц не проблема для SQLite).
"""
import sqlite3
import datetime
from pathlib import Path
from contextlib import contextmanager

DB_PATH = Path(__file__).parent / "data" / "boris.db"


def week_start(d: datetime.date) -> datetime.date:
    """Понедельник той недели, в которую попадает дата d."""
    return d - datetime.timedelta(days=d.weekday())


@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                platform TEXT NOT NULL,          -- 'ozon' | 'wb'
                tx_date TEXT NOT NULL,           -- ISO дата операции
                week_start TEXT NOT NULL,        -- ISO дата понедельника недели
                group_name TEXT NOT NULL,        -- статья PnL (Продажи, Логистика, ...)
                article TEXT,                    -- код товара (для себестоимости)
                qty REAL DEFAULT 0,
                amount REAL NOT NULL,            -- сумма по статье (со знаком)
                gross_amount REAL DEFAULT 0,      -- для ВБ: заказ до вычета комиссии
                source_file TEXT,
                uploaded_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_tx_week ON transactions(week_start);
            CREATE INDEX IF NOT EXISTS idx_tx_platform ON transactions(platform);

            CREATE TABLE IF NOT EXISTS product_costs (
                article TEXT PRIMARY KEY,
                name TEXT,
                cost_per_unit REAL
            );

            CREATE TABLE IF NOT EXISTS processed_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                platform TEXT NOT NULL,
                period_start TEXT,
                period_end TEXT,
                rows_added INTEGER,
                status TEXT,                     -- 'ok' | 'duplicate' | 'error'
                message TEXT,
                uploaded_at TEXT NOT NULL
            );
            """
        )


def file_already_processed(filename: str, platform: str) -> bool:
    with get_conn() as conn:
        row = conn.execute(
            "SELECT 1 FROM processed_files WHERE filename=? AND platform=? AND status='ok'",
            (filename, platform),
        ).fetchone()
        return row is not None


def log_processed_file(filename, platform, period_start, period_end, rows_added, status, message=""):
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO processed_files
               (filename, platform, period_start, period_end, rows_added, status, message, uploaded_at)
               VALUES (?,?,?,?,?,?,?,?)""",
            (filename, platform, str(period_start), str(period_end), rows_added, status, message,
             datetime.datetime.now().isoformat()),
        )


def insert_transactions(rows):
    """rows: список dict с ключами platform, tx_date, group_name, article, qty, amount,
    gross_amount, source_file"""
    now = datetime.datetime.now().isoformat()
    with get_conn() as conn:
        conn.executemany(
            """INSERT INTO transactions
               (platform, tx_date, week_start, group_name, article, qty, amount, gross_amount,
                source_file, uploaded_at)
               VALUES (:platform,:tx_date,:week_start,:group_name,:article,:qty,:amount,
                       :gross_amount,:source_file,:uploaded_at)""",
            [
                {**r, "week_start": str(week_start(datetime.date.fromisoformat(r["tx_date"]))),
                 "uploaded_at": now}
                for r in rows
            ],
        )


def delete_transactions_for_file(filename: str):
    with get_conn() as conn:
        conn.execute("DELETE FROM transactions WHERE source_file=?", (filename,))


def get_all_week_starts():
    with get_conn() as conn:
        rows = conn.execute("SELECT DISTINCT week_start FROM transactions ORDER BY week_start").fetchall()
        return [datetime.date.fromisoformat(r["week_start"]) for r in rows]


def get_week_summary(week_start_date: datetime.date):
    """Возвращает суммы по каждой статье для Ozon и ВБ за неделю (для сборки текстовой сводки и Excel)."""
    ws = str(week_start_date)
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT platform, group_name, SUM(amount) as total FROM transactions "
            "WHERE week_start=? GROUP BY platform, group_name",
            (ws,),
        ).fetchall()
        gross = conn.execute(
            "SELECT platform, SUM(gross_amount) as total FROM transactions "
            "WHERE week_start=? AND group_name='Продажи' GROUP BY platform",
            (ws,),
        ).fetchall()
        cogs = conn.execute(
            """SELECT t.platform, SUM(t.qty * COALESCE(p.cost_per_unit,0)) as cogs
               FROM transactions t LEFT JOIN product_costs p ON t.article = p.article
               WHERE t.week_start=? AND t.group_name='Продажи'
               GROUP BY t.platform""",
            (ws,),
        ).fetchall()
    return {
        "by_group": {(r["platform"], r["group_name"]): r["total"] for r in rows},
        "gross": {r["platform"]: r["total"] for r in gross},
        "cogs": {r["platform"]: r["cogs"] for r in cogs},
    }


def upsert_product_cost(article: str, name: str, cost_per_unit: float):
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO product_costs (article, name, cost_per_unit) VALUES (?,?,?)
               ON CONFLICT(article) DO UPDATE SET name=excluded.name, cost_per_unit=excluded.cost_per_unit""",
            (article, name, cost_per_unit),
        )


def get_missing_cost_articles():
    """Товары, которые продавались, но себестоимость для них не указана."""
    with get_conn() as conn:
        rows = conn.execute(
            """SELECT DISTINCT t.article, MAX(t.qty) as qty
               FROM transactions t LEFT JOIN product_costs p ON t.article = p.article
               WHERE t.group_name='Продажи' AND t.article IS NOT NULL
                 AND (p.cost_per_unit IS NULL)
               GROUP BY t.article"""
        ).fetchall()
        return [r["article"] for r in rows]
