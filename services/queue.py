"""
Простая фоновая очередь обработки файлов на asyncio.

Задача кладётся в очередь сразу после получения файла от пользователя —
бот тут же отвечает "приняла" и не блокируется. Файлы разбираются
по одному в отдельной asyncio-задаче (worker), не мешая боту отвечать
на другие сообщения.
"""
import asyncio
import datetime
import logging

import db
from parsers.detect import detect_report_type
from parsers import ozon as parse_ozon
from parsers import wb as parse_wb
from parsers.ozon import parse_ozon_report, get_period as get_period_ozon
from parsers.wb import parse_wb_report, get_period as get_period_wb

log = logging.getLogger("boris.queue")

task_queue: asyncio.Queue = asyncio.Queue()


class UploadJob:
    def __init__(self, chat_id, file_path, filename, status_message_id=None):
        self.chat_id = chat_id
        self.file_path = file_path
        self.filename = filename
        self.status_message_id = status_message_id


async def enqueue(job: UploadJob):
    await task_queue.put(job)


def _process_one(job: UploadJob) -> dict:
    """Синхронная обработка одного файла (парсинг + запись в БД).
    Выполняется в отдельном потоке через asyncio.to_thread, чтобы не блокировать event loop."""
    kind = detect_report_type(job.file_path)
    if kind == "unknown":
        db.log_skipped_file(job.filename, "unknown", None, "error", "не распознан тип файла")
        return {"status": "error", "message": f"«{job.filename}» — не похож на отчёт Ozon или ВБ, пропускаю."}

    if kind == "ozon":
        period = get_period_ozon(job.file_path)
        rows = parse_ozon_report(job.file_path, job.filename)
        chash = parse_ozon.content_hash(rows) if hasattr(parse_ozon, "content_hash") else None
    else:
        period = get_period_wb(job.file_path)
        rows = parse_wb_report(job.file_path, job.filename)
        chash = parse_wb.content_hash(rows) if hasattr(parse_wb, "content_hash") else None

    # проверка по содержимому — ловит повторную загрузку того же отчёта
    # даже если файл называется по-другому (частый случай при переcкачивании)
    if chash:
        prev_filename = db.find_by_content_hash(chash, kind)
        if prev_filename:
            db.log_skipped_file(job.filename, kind, period, "duplicate", f"дубликат {prev_filename}")
            return {"status": "duplicate",
                    "message": f"«{job.filename}» — это те же данные, что уже были загружены "
                               f"ранее как «{prev_filename}», пропускаю (не задваиваю)."}

    file_id = db.insert_report(rows, job.filename, kind, period, chash)

    return {
        "status": "ok",
        "platform": kind,
        "period": period,
        "rows": len(rows),
        "file_id": file_id,
        "message": f"«{job.filename}» — {('Ozon' if kind=='ozon' else 'ВБ')}, "
                   f"{period[0].strftime('%d.%m.%Y')}–{period[1].strftime('%d.%m.%Y')}, "
                   f"строк: {len(rows)} (№{file_id} в базе — этим номером можно удалить через /delete)",
    }


async def worker(bot):
    """Бесконечный воркер: берёт задачи из очереди и обрабатывает по одной."""
    while True:
        job: UploadJob = await task_queue.get()
        try:
            result = await asyncio.to_thread(_process_one, job)
            await bot.send_message(job.chat_id, result["message"])
        except Exception as e:
            log.exception("Ошибка обработки файла %s", job.filename)
            await bot.send_message(job.chat_id, f"❌ Ошибка при обработке «{job.filename}»: {e}")
        finally:
            task_queue.task_done()


async def process_batch(bot, chat_id, jobs: list[UploadJob]):
    """Обрабатывает пачку файлов последовательно, присылая прогресс и финальную сводку."""
    await bot.send_message(chat_id, f"Приняла {len(jobs)} файл(ов), начинаю обработку в фоне…")

    results = []
    for i, job in enumerate(jobs, start=1):
        result = await asyncio.to_thread(_process_one, job)
        results.append(result)
        if i % 5 == 0 or i == len(jobs):
            await bot.send_message(chat_id, f"Обработано {i} из {len(jobs)}…")

    ok = [r for r in results if r["status"] == "ok"]
    dup = [r for r in results if r["status"] == "duplicate"]
    err = [r for r in results if r["status"] == "error"]

    lines = [f"✅ Готово! Успешно обработано: {len(ok)} из {len(jobs)}"]
    if dup:
        lines.append(f"⏭ Пропущено (уже загружено ранее): {len(dup)}")
    if err:
        lines.append(f"⚠️ Не удалось обработать: {len(err)}")
        for r in err:
            lines.append(f"   • {r['message']}")

    weeks = db.get_all_week_starts()
    if weeks:
        lines.append(f"\nВ базе данные за {len(weeks)} недель(и): "
                      f"{weeks[0].strftime('%d.%m.%Y')} – {(weeks[-1] + datetime.timedelta(days=6)).strftime('%d.%m.%Y')}")

    await bot.send_message(chat_id, "\n".join(lines))
