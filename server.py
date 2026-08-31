"""
Борис — бот, который собирает PnL по отчётам Ozon и Wildberries.

Всё управление — через кнопки в меню снизу. Команды тоже работают (на
случай, если кнопки не видно), но ими можно не пользоваться:
  /start      — показать меню
  /export     — выгрузить Excel-файл со всем накопленным PnL
  /week       — сводка по неделе (список недель — кнопками)
  /cost       — товары без указанной себестоимости
  /setcost    — задать себестоимость одного товара вручную (только текстом)
  /importcost — загрузить себестоимость файлом
  /files      — список загруженных отчётов (удаление — кнопкой)
  /delete     — удалить конкретный отчёт по номеру
  /status     — сводный статус базы

Загрузка отчёта: просто прикрепить .xlsx (Ozon) или .zip (ВБ) файлом в чат —
Борис сам определит тип, период и добавит данные в базу. Несколько файлов
можно отправить одним сообщением или подряд — они уйдут в фоновую очередь
и обработаются один за другим, не блокируя бота. Если тот же отчёт (даже
под другим именем файла) уже был загружен раньше — Борис это заметит и
не задвоит данные.
"""
import asyncio
import logging
import os
import sys
import tempfile
import datetime

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import (
    Message, FSInputFile, CallbackQuery,
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton,
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

import db
from services import export as export_service
from services.queue import UploadJob, process_batch
from parsers.cost import parse_cost_file

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
dp = Dispatcher(storage=MemoryStorage())


class ImportStates(StatesGroup):
    waiting_cost_file = State()


# буфер: если пользователь прислал несколько файлов подряд (альбомом или быстро один
# за другим), собираем их в одну пачку перед обработкой, чтобы не плодить отдельные
# статусные сообщения на каждый файл
_pending_files: dict[int, list] = {}
_pending_timers: dict[int, asyncio.Task] = {}
BATCH_DELAY_SECONDS = 2.5

# ----------------------------- Меню (кнопки) -----------------------------

BTN_WEEK = "📊 Сводка за неделю"
BTN_FILES = "📁 Мои отчёты"
BTN_COST = "💰 Себестоимость"
BTN_EXPORT = "📤 Выгрузить PnL"
BTN_STATUS = "ℹ️ Статус"

main_menu = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text=BTN_WEEK), KeyboardButton(text=BTN_FILES)],
        [KeyboardButton(text=BTN_COST), KeyboardButton(text=BTN_EXPORT)],
        [KeyboardButton(text=BTN_STATUS)],
    ],
    resize_keyboard=True,
)


def weeks_keyboard(weeks: list[datetime.date]) -> InlineKeyboardMarkup:
    """Кнопки с недельными периодами — последние сверху, не больше 10 штук."""
    rows = []
    for w in list(reversed(weeks))[:10]:
        end = w + datetime.timedelta(days=6)
        label = f"{w.strftime('%d.%m')}–{end.strftime('%d.%m.%Y')}"
        rows.append([InlineKeyboardButton(text=label, callback_data=f"week:{w.isoformat()}")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def files_keyboard(files: list[dict]) -> InlineKeyboardMarkup:
    """По кнопке 'Удалить' на каждый ещё не удалённый отчёт."""
    rows = []
    for f in files:
        if f["status"] != "ok":
            continue
        label = f"🗑 Удалить №{f['id']} ({f['platform']}, {f['period_start']}–{f['period_end']})"
        rows.append([InlineKeyboardButton(text=label, callback_data=f"del:{f['id']}")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def confirm_delete_keyboard(file_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Да, удалить", callback_data=f"delyes:{file_id}"),
        InlineKeyboardButton(text="✖️ Отмена", callback_data=f"delno:{file_id}"),
    ]])


def cost_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="📥 Загрузить файлом", callback_data="importcost"),
    ]])


# ----------------------------- Общая логика (переиспользуется командами и кнопками) -----------------------------

async def send_status(message: Message):
    weeks = db.get_all_week_starts()
    if not weeks:
        await message.answer("Пока данных нет — пришли первый отчёт (.xlsx Ozon или .zip ВБ).", reply_markup=main_menu)
        return
    first, last = weeks[0], weeks[-1] + datetime.timedelta(days=6)
    missing = db.get_missing_cost_articles()
    text = f"В базе {len(weeks)} недель(и): {first.strftime('%d.%m.%Y')} – {last.strftime('%d.%m.%Y')}\n"
    if missing:
        text += f"⚠️ Без себестоимости: {len(missing)} товар(ов)"
    await message.answer(text, reply_markup=main_menu)


async def send_week_menu(message: Message):
    weeks = db.get_all_week_starts()
    if not weeks:
        await message.answer("Пока данных нет — пришли первый отчёт.", reply_markup=main_menu)
        return
    await message.answer("За какую неделю показать сводку?", reply_markup=weeks_keyboard(weeks))


async def send_files_list(message: Message):
    files = db.list_files(limit=20)
    if not files:
        await message.answer("Пока ничего не загружено.", reply_markup=main_menu)
        return
    lines = ["Последние загруженные отчёты (новые сверху):\n"]
    for f in files:
        status_icon = {"ok": "✅", "duplicate": "⏭", "error": "⚠️", "deleted": "🗑"}.get(f["status"], "•")
        period = f"{f['period_start']}–{f['period_end']}" if f["period_start"] else "—"
        lines.append(f"{status_icon} №{f['id']} · {f['platform']} · {period} · {f['rows_added']} строк · {f['filename']}")
    await message.answer("\n".join(lines), reply_markup=files_keyboard(files))


async def send_cost_status(message: Message):
    missing = db.get_missing_cost_articles()
    if not missing:
        await message.answer("Себестоимость указана для всех продающихся товаров ✅", reply_markup=main_menu)
        return
    lines = ["Товары без указанной себестоимости:\n"]
    for art in missing[:30]:
        lines.append(f"• {art}")
    if len(missing) > 30:
        lines.append(f"…и ещё {len(missing) - 30}")
    lines.append("\nЗадать цену вручную: /setcost <код> <цена>\nНапример: /setcost 36095832 150")
    await message.answer("\n".join(lines), reply_markup=cost_keyboard())


async def send_export(message: Message):
    weeks = db.get_all_week_starts()
    if not weeks:
        await message.answer("Пока данных нет — пришли первый отчёт.", reply_markup=main_menu)
        return
    await message.answer("Собираю файл…")
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "PnL_Ozon_WB_общий.xlsx")
        export_service.build_full_workbook(path)
        await message.answer_document(FSInputFile(path), reply_markup=main_menu)


async def start_importcost(message_or_callback, state: FSMContext):
    await state.set_state(ImportStates.waiting_cost_file)
    text = (
        "Пришли файл с себестоимостью (.xlsx или .csv).\n\n"
        "Нужны две колонки: код товара («Код»/«Артикул»/«SKU») и цена "
        "(«Себестоимость»/«Цена»/«Стоимость»). Названия колонок могут быть "
        "почти любыми — я найду их сама. Необязательная колонка «Название» "
        "тоже подхватится, если есть.\n\n"
        "Если товар с таким кодом уже есть — его цена обновится."
    )
    if isinstance(message_or_callback, CallbackQuery):
        await message_or_callback.message.answer(text)
    else:
        await message_or_callback.answer(text)


# ----------------------------- Команды -----------------------------

@dp.message(Command("start"))
async def cmd_start(message: Message):
    await message.answer(
        "Привет! Я Борис 🤖\n\n"
        "Пришли мне отчёт по начислениям Ozon (.xlsx) или еженедельный "
        "детализированный отчёт Wildberries (.zip) — я сам разберу файл, "
        "определю неделю по датам внутри него и соберу общий PnL.\n\n"
        "Можно прислать сразу много файлов за раз — я обработаю их по очереди в фоне.\n\n"
        "Дальше пользуйся кнопками внизу 👇",
        reply_markup=main_menu,
    )


@dp.message(Command("status"))
async def cmd_status(message: Message):
    await send_status(message)


@dp.message(Command("week"))
async def cmd_week(message: Message):
    weeks = db.get_all_week_starts()
    if not weeks:
        await message.answer("Пока данных нет — пришли первый отчёт.", reply_markup=main_menu)
        return
    parts = message.text.split()
    if len(parts) > 1:
        try:
            requested = datetime.date.fromisoformat(parts[1])
            target = requested - datetime.timedelta(days=requested.weekday())
        except ValueError:
            await message.answer("Не поняла дату. Формат: /week 2026-07-20")
            return
        await message.answer(export_service.build_week_summary_text(target), reply_markup=main_menu)
    else:
        await send_week_menu(message)


@dp.message(Command("cost"))
async def cmd_cost(message: Message):
    await send_cost_status(message)


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
    await message.answer(f"Себестоимость для «{article}» установлена: {price:.0f} ₽/шт.", reply_markup=main_menu)


@dp.message(Command("importcost"))
async def cmd_importcost(message: Message, state: FSMContext):
    await start_importcost(message, state)


@dp.message(Command("files"))
async def cmd_files(message: Message):
    await send_files_list(message)


@dp.message(Command("delete"))
async def cmd_delete(message: Message):
    parts = message.text.split()
    if len(parts) != 2 or not parts[1].isdigit():
        await message.answer("Формат: /delete <номер>\nНомер смотри в /files")
        return
    file_id = int(parts[1])
    file_row = db.get_file(file_id)
    if not file_row:
        await message.answer(f"Не нашла отчёт №{file_id}.")
        return
    if file_row["status"] == "deleted":
        await message.answer(f"Отчёт №{file_id} уже был удалён раньше.")
        return
    deleted_rows = db.delete_file(file_id)
    await message.answer(
        f"Удалила отчёт №{file_id} («{file_row['filename']}») — {deleted_rows} строк убрано из базы.",
        reply_markup=main_menu,
    )


@dp.message(Command("export"))
async def cmd_export(message: Message):
    await send_export(message)


# ----------------------------- Кнопки главного меню (обычный текст) -----------------------------

@dp.message(F.text == BTN_WEEK)
async def btn_week(message: Message):
    await send_week_menu(message)


@dp.message(F.text == BTN_FILES)
async def btn_files(message: Message):
    await send_files_list(message)


@dp.message(F.text == BTN_COST)
async def btn_cost(message: Message):
    await send_cost_status(message)


@dp.message(F.text == BTN_EXPORT)
async def btn_export(message: Message):
    await send_export(message)


@dp.message(F.text == BTN_STATUS)
async def btn_status(message: Message):
    await send_status(message)


# ----------------------------- Инлайн-кнопки (callback) -----------------------------

@dp.callback_query(F.data.startswith("week:"))
async def cb_week(callback: CallbackQuery):
    date_str = callback.data.split(":", 1)[1]
    target = datetime.date.fromisoformat(date_str)
    text = export_service.build_week_summary_text(target)
    await callback.message.answer(text, reply_markup=main_menu)
    await callback.answer()


@dp.callback_query(F.data.startswith("del:"))
async def cb_delete_ask(callback: CallbackQuery):
    file_id = int(callback.data.split(":", 1)[1])
    file_row = db.get_file(file_id)
    if not file_row or file_row["status"] != "ok":
        await callback.answer("Этот отчёт уже недоступен для удаления.", show_alert=True)
        return
    await callback.message.answer(
        f"Точно удалить отчёт №{file_id} («{file_row['filename']}», {file_row['rows_added']} строк)?",
        reply_markup=confirm_delete_keyboard(file_id),
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("delyes:"))
async def cb_delete_confirm(callback: CallbackQuery):
    file_id = int(callback.data.split(":", 1)[1])
    file_row = db.get_file(file_id)
    if not file_row or file_row["status"] != "ok":
        await callback.message.edit_text("Этот отчёт уже был удалён раньше.")
        await callback.answer()
        return
    deleted_rows = db.delete_file(file_id)
    await callback.message.edit_text(
        f"✅ Удалила отчёт №{file_id} («{file_row['filename']}») — {deleted_rows} строк убрано из базы."
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("delno:"))
async def cb_delete_cancel(callback: CallbackQuery):
    await callback.message.edit_text("Отменила, ничего не удалила.")
    await callback.answer()


@dp.callback_query(F.data == "importcost")
async def cb_importcost(callback: CallbackQuery, state: FSMContext):
    await start_importcost(callback, state)
    await callback.answer()


# ----------------------------- Загрузка файлов -----------------------------

async def _flush_pending(chat_id: int):
    await asyncio.sleep(BATCH_DELAY_SECONDS)
    jobs = _pending_files.pop(chat_id, [])
    _pending_timers.pop(chat_id, None)
    if jobs:
        await process_batch(bot, chat_id, jobs)


@dp.message(F.document)
async def handle_document(message: Message, state: FSMContext):
    doc = message.document
    filename = doc.file_name or f"file_{doc.file_id}"

    current_state = await state.get_state()
    if current_state == ImportStates.waiting_cost_file.state:
        await state.clear()
        tmpdir = tempfile.mkdtemp()
        local_path = os.path.join(tmpdir, filename)
        file_info = await bot.get_file(doc.file_id)
        await bot.download_file(file_info.file_path, destination=local_path)
        try:
            items = parse_cost_file(local_path)
        except Exception as e:
            await message.answer(f"Не смогла прочитать файл: {e}", reply_markup=main_menu)
            return
        db.upsert_product_costs_bulk(items)
        await message.answer(f"Загрузила себестоимость для {len(items)} товар(ов) ✅", reply_markup=main_menu)
        return

    if not (filename.lower().endswith(".xlsx") or filename.lower().endswith(".zip")):
        await message.answer(f"«{filename}» — не похоже на отчёт Ozon (.xlsx) или ВБ (.zip), пропускаю.")
        return

    tmpdir = tempfile.mkdtemp()
    local_path = os.path.join(tmpdir, filename)
    file_info = await bot.get_file(doc.file_id)
    await bot.download_file(file_info.file_path, destination=local_path)

    job = UploadJob(chat_id=message.chat.id, file_path=local_path, filename=filename)
    _pending_files.setdefault(message.chat.id, []).append(job)

    old_timer = _pending_timers.get(message.chat.id)
    if old_timer:
        old_timer.cancel()
    _pending_timers[message.chat.id] = asyncio.create_task(_flush_pending(message.chat.id))


async def main():
    db.init_db()
    log.info("Борис запущен")
    await bot.delete_webhook(drop_pending_updates=False)
    await dp.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception:
        log.exception("Борис упал при запуске:")
        raise
