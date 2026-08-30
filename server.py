"""
Борис — бот, который собирает PnL по отчётам Ozon и Wildberries.

Команды:
  /start   — приветствие и краткая инструкция
  /export  — прислать актуальный Excel-файл со всем накопленным PnL
  /week    — прислать текстовую сводку по конкретной неделе (последней по умолчанию)
  /cost    — показать товары без указанной себестоимости / задать цену
  /status  — сколько недель и файлов уже в базе

Загрузка отчёта: просто прикрепить .xlsx (Ozon) или .zip (ВБ) файлом в чат —
Борис сам определит тип, период и добавит данные в базу. Несколько файлов
можно отправить одним сообщением или подряд — они уйдут в фоновую очередь
и обработаются один за другим, не блокируя бота.
"""
import asyncio
import logging
import os
import sys
import tempfile
import datetime

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile

import db
from services import export as export_service
from services.queue import UploadJob, process_batch

# Явно пишем логи в stdout — многие хостинги (в т.ч. bothost) показывают
# в панели "Логи" именно stdout, а не stderr (куда logging пишет по умолчанию).
logging.basicConfig(
    level=logging.INFO,
    stream=sys.stdout,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("boris")

# Поддерживаем несколько распространённых названий переменной с токеном —
# на разных хостингах (и в разных шаблонах bothost) поле может называться по-разному.
_TOKEN_ENV_NAMES = ["BOT_TOKEN", "BORIS_BOT_TOKEN", "TOKEN", "TELEGRAM_TOKEN", "TELEGRAM_BOT_TOKEN"]
BOT_TOKEN = ""
for _name in _TOKEN_ENV_NAMES:
    _val = os.environ.get(_name, "").strip()
    if _val:
        BOT_TOKEN = _val
        break

if not BOT_TOKEN:
    print(
        "ОШИБКА: не нашла токен бота ни в одной из переменных окружения: "
        + ", ".join(_TOKEN_ENV_NAMES),
        flush=True,
    )
    print("Проверь, что токен вписан в одну из них в настройках проекта на bothost.", flush=True)
    sys.exit(1)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# буфер: если пользователь прислал несколько файлов подряд (альбомом или быстро один
# за другим), собираем их в одну пачку перед обработкой, чтобы не плодить отдельные
# статусные сообщения на каждый файл
_pending_files: dict[int, list] = {}
_pending_timers: dict[int, asyncio.Task] = {}
BATCH_DELAY_SECONDS = 2.5


@dp.message(Command("start"))
async def cmd_start(message: Message):
    await message.answer(
        "Привет! Я Борис 🤖\n\n"
        "Пришли мне отчёт по начислениям Ozon (.xlsx) или еженедельный "
        "детализированный отчёт Wildberries (.zip) — я сам разберу файл, "
        "определю неделю по датам внутри него и соберу общий PnL.\n\n"
        "Можно прислать сразу много файлов за раз — я обработаю их по очереди в фоне.\n\n"
        "Команды:\n"
        "/export — выгрузить весь накопленный PnL Excel-файлом\n"
        "/week — сводка за последнюю неделю с данными\n"
        "/cost — товары без указанной себестоимости\n"
        "/status — что уже есть в базе"
    )


@dp.message(Command("status"))
async def cmd_status(message: Message):
    weeks = db.get_all_week_starts()
    if not weeks:
        await message.answer("Пока данных нет — пришли первый отчёт.")
        return
    first, last = weeks[0], weeks[-1] + datetime.timedelta(days=6)
    missing = db.get_missing_cost_articles()
    text = (
        f"В базе {len(weeks)} недель(и): {first.strftime('%d.%m.%Y')} – {last.strftime('%d.%m.%Y')}\n"
    )
    if missing:
        text += f"⚠️ Без себестоимости: {len(missing)} товар(ов) — /cost"
    await message.answer(text)


@dp.message(Command("week"))
async def cmd_week(message: Message):
    weeks = db.get_all_week_starts()
    if not weeks:
        await message.answer("Пока данных нет — пришли первый отчёт.")
        return
    # /week 2026-07-20 — конкретная неделя (по понедельнику), иначе — последняя доступная
    parts = message.text.split()
    target = weeks[-1]
    if len(parts) > 1:
        try:
            requested = datetime.date.fromisoformat(parts[1])
            target = requested - datetime.timedelta(days=requested.weekday())
        except ValueError:
            await message.answer("Не поняла дату. Формат: /week 2026-07-20")
            return
    text = export_service.build_week_summary_text(target)
    await message.answer(text)


@dp.message(Command("cost"))
async def cmd_cost(message: Message):
    missing = db.get_missing_cost_articles()
    if not missing:
        await message.answer("Себестоимость указана для всех продающихся товаров ✅")
        return
    lines = ["Товары без указанной себестоимости:\n"]
    for art in missing[:30]:
        lines.append(f"• {art}")
    if len(missing) > 30:
        lines.append(f"…и ещё {len(missing) - 30}")
    lines.append(
        "\nЧтобы указать цену, напиши: /setcost <код> <цена>\n"
        "Например: /setcost 36095832 150"
    )
    await message.answer("\n".join(lines))


@dp.message(Command("setcost"))
async def cmd_setcost(message: Message):
    parts = message.text.split()
    if len(parts) != 3:
        await message.answer("Формат: /setcost <код> <цена>\nНапример: /setcost 36095832 150")
        return
    _, article, price_str = parts
    try:
        price = float(price_str.replace(",", "."))
    except ValueError:
        await message.answer("Цена должна быть числом.")
        return
    db.upsert_product_cost(article, None, price)
    await message.answer(f"Себестоимость для «{article}» установлена: {price:.0f} ₽/шт.")


@dp.message(Command("export"))
async def cmd_export(message: Message):
    weeks = db.get_all_week_starts()
    if not weeks:
        await message.answer("Пока данных нет — пришли первый отчёт.")
        return
    await message.answer("Собираю файл…")
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "PnL_Ozon_WB_общий.xlsx")
        export_service.build_full_workbook(path)
        await message.answer_document(FSInputFile(path))


async def _flush_pending(chat_id: int):
    await asyncio.sleep(BATCH_DELAY_SECONDS)
    jobs = _pending_files.pop(chat_id, [])
    _pending_timers.pop(chat_id, None)
    if jobs:
        await process_batch(bot, chat_id, jobs)


@dp.message(F.document)
async def handle_document(message: Message):
    doc = message.document
    filename = doc.file_name or f"file_{doc.file_id}"
    if not (filename.lower().endswith(".xlsx") or filename.lower().endswith(".zip")):
        await message.answer(
            f"«{filename}» — не похоже на отчёт Ozon (.xlsx) или ВБ (.zip), пропускаю."
        )
        return

    tmpdir = tempfile.mkdtemp()
    local_path = os.path.join(tmpdir, filename)
    file_info = await bot.get_file(doc.file_id)
    await bot.download_file(file_info.file_path, destination=local_path)

    job = UploadJob(chat_id=message.chat.id, file_path=local_path, filename=filename)
    _pending_files.setdefault(message.chat.id, []).append(job)

    # сбрасываем таймер батча — ждём ещё немного, вдруг пришлют ещё файлы следом
    old_timer = _pending_timers.get(message.chat.id)
    if old_timer:
        old_timer.cancel()
    _pending_timers[message.chat.id] = asyncio.create_task(_flush_pending(message.chat.id))


async def main():
    db.init_db()
    log.info("Борис запущен")
    # Сбрасываем возможный вебхук — если он был поставлен раньше (например,
    # другим способом запуска), long polling не будет получать сообщения,
    # пока вебхук не снят.
    await bot.delete_webhook(drop_pending_updates=False)
    await dp.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception:
        log.exception("Борис упал при запуске:")
        raise
