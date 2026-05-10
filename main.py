"""
bot.py — Swimming Training Bot
"""

import asyncio
import os
import time
from datetime import datetime
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, CallbackQuery, ReplyKeyboardMarkup, KeyboardButton, FSInputFile
from aiogram.filters import CommandStart
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.exceptions import TelegramBadRequest

load_dotenv()

# ═══════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════

TOKEN          = os.getenv("BOT_TOKEN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
TRAINER_ID     = int(os.getenv("TRAINER_ID", "0"))

USER_COLUMNS = {
    "Игорь":    "D",
    "Марк":     "E",
    "Аня":      "F",
    "Матвей":   "G",
    "Ульяна":   "H",
    "Руслан":   "I",
    "Вероника": "J",
    "Виталик":  "K",
    "Антон":    "L",
    "Женя":     "M",
    "Настя":    "N",
    "Савелий":  "O",
    "Маша":     "P",
}

# ═══════════════════════════════════════════════════════════════
# GOOGLE SHEETS
# ═══════════════════════════════════════════════════════════════

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
gc    = gspread.authorize(creds)
ss    = gc.open_by_key(SPREADSHEET_ID)

def get_bd_sheet():
    return ss.worksheet("BD")

def get_source_sheet():
    return ss.sheet1

# ═══════════════════════════════════════════════════════════════
# КЭШ
# ═══════════════════════════════════════════════════════════════

_user_cache:          dict[int, str] = {}   # telegram_id → имя
_tid_cache:           dict[str, int] = {}   # имя → telegram_id
_bd_rows:             list           = []   # строки из листа BD
_bd_ts:               float          = 0.0  # время последней синхронизации
_week_marker_row:       int  = -1    # строка с "текущая неделя"
_week_session_notified: bool = True  # True = уже уведомили (на старте не спамим)
_last_booking:        dict[str, str] = {}   # имя → дата последнего бронирования "DD.MM.YYYY"
_inactivity_notified: set[str]       = set() # кому уже отправили "давно не виделись"
_reminded_training:   set[str]       = set() # "имя|дата" — кому уже напомнили о тренировке
_reminder_date:       str            = ""    # дата последней проверки напоминаний (сброс в полночь)
_known_plans:         dict[str, str] = {}    # "имя|дата" → план (для уведомления о новом плане)
BD_TTL = 300                                 # секунд (5 минут)

def _invalidate_bd():
    global _bd_ts
    _bd_ts = 0.0

def load_all_tids():
    """Загружает telegram_id всех пользователей из строки 6 таблицы в кэш."""
    try:
        row = get_source_sheet().row_values(6)
        for name, col_letter in USER_COLUMNS.items():
            col_idx = ord(col_letter) - ord("A")
            if col_idx < len(row):
                val = str(row[col_idx]).strip()
                if val and val not in ("", "0", "None"):
                    try:
                        tid = int(val)
                        _tid_cache[name]  = tid
                        _user_cache[tid]  = name
                    except ValueError:
                        pass
    except Exception as e:
        print(f"load_all_tids ошибка: {e}")

def _ensure_bd():
    global _bd_rows, _bd_ts
    if (time.time() - _bd_ts) > BD_TTL:
        sync_bd()
        _bd_rows = get_bd_sheet().get_all_values()
        _bd_ts   = time.time()

# ═══════════════════════════════════════════════════════════════
# ПОЛЬЗОВАТЕЛИ
# ═══════════════════════════════════════════════════════════════

def get_user_name_by_telegram_id(telegram_id: int) -> str | None:
    if telegram_id in _user_cache:
        return _user_cache[telegram_id]
    data = get_source_sheet().get_all_values()
    if len(data) < 6:
        return None
    row = data[5]
    for name, col_letter in USER_COLUMNS.items():
        col_idx = ord(col_letter) - ord("A")
        if col_idx < len(row) and str(row[col_idx]).strip() == str(telegram_id):
            _user_cache[telegram_id] = name
            _tid_cache[name] = telegram_id
            return name
    return None

def save_telegram_id(user_name: str, telegram_id: int):
    col = USER_COLUMNS[user_name]
    get_source_sheet().update(f"{col}6", [[telegram_id]])
    _user_cache[telegram_id] = user_name
    _tid_cache[user_name] = telegram_id

def save_last_booking(user_name: str, date_str: str):
    col = USER_COLUMNS[user_name]
    try:
        get_source_sheet().update(f"{col}7", [[date_str]])
    except Exception as e:
        print(f"save_last_booking({user_name}) ошибка: {e}")

def load_last_bookings():
    """Загружает даты последних бронирований из строки 7 таблицы."""
    try:
        row = get_source_sheet().row_values(7)
        for name, col_letter in USER_COLUMNS.items():
            col_idx = ord(col_letter) - ord("A")
            if col_idx < len(row):
                val = str(row[col_idx]).strip()
                if val and val not in ("", "None"):
                    _last_booking[name] = val
    except Exception as e:
        print(f"load_last_bookings ошибка: {e}")

# ═══════════════════════════════════════════════════════════════
# РАСПИСАНИЕ ИЗ BD
# ═══════════════════════════════════════════════════════════════

DAY_RU = {
    "Monday": "Понедельник", "Tuesday": "Вторник",  "Wednesday": "Среда",
    "Thursday": "Четверг",   "Friday":  "Пятница",  "Saturday":  "Суббота",
    "Sunday": "Воскресенье",
}

DAY_SHORT = {
    "Понедельник": "Пн", "Вторник": "Вт", "Среда":       "Ср",
    "Четверг":     "Чт", "Пятница": "Пт", "Суббота":     "Сб",
    "Воскресенье": "Вс",
}

def get_week_info() -> dict:
    cell = _bd_rows[0][0] if _bd_rows and _bd_rows[0] else ""
    parts = [p.strip() for p in cell.split("|")]
    return {
        "label": parts[0] if len(parts) > 0 else "",
        "dates": parts[1] if len(parts) > 1 else "",
    }

def get_schedule_for_user(user_name: str) -> list[dict]:
    if len(_bd_rows) < 2:
        return []
    headers = _bd_rows[1]
    col = {h: i for i, h in enumerate(headers)}
    result = []
    for row in _bd_rows[2:]:
        if not row or len(row) <= col.get("User", 0):
            continue
        if row[col["User"]].strip() != user_name:
            continue
        result.append({
            "day":      DAY_RU.get(row[col.get("Day", 2)], row[col.get("Day", 2)]),
            "date":     row[col.get("Date", 1)],
            "time":     row[col.get("Time", 3)],
            "plan":     row[col.get("Plan", 4)],
            "volume":   row[col.get("Volume", 5)],
            "remote":   row[col.get("Remote", 6)].lower() == "yes",
            "booked":   row[col.get("Booked", 7)].upper() == "TRUE",
            "comments": row[col.get("Comments", 8)],
        })
    return result


# ═══════════════════════════════════════════════════════════════
# СИНХРОНИЗАЦИЯ BD
# ═══════════════════════════════════════════════════════════════

DAY_MAP_SYNC = {
    "понедельник": "Monday", "вторник": "Tuesday",  "среда":       "Wednesday",
    "четверг":     "Thursday", "пятница": "Friday", "суббота":     "Saturday",
    "воскресенье": "Sunday",
}

USER_COLS_SYNC = list(range(3, 26))  # колонки D..Z (0-based)

def sync_bd():
    """
    Читает лист 2026, находит текущую неделю, формирует плоскую таблицу
    и перезаписывает лист BD. Вызывается перед каждым показом расписания/записей.
    """
    src  = get_source_sheet()
    data = src.get_all_values()

    # Находим строку с маркером текущей недели
    week_bron_idx = None
    for i, row in enumerate(data):
        last = str(row[-1]).strip().lower() if row else ""
        if "текущ" in last:
            week_bron_idx = i
            break
    if week_bron_idx is None:
        get_bd_sheet().clear()  # сбрасываем BD чтобы кэш не показывал старые данные
        return

    # Заголовок недели (строка выше блока Забронировать)
    week_label = ""
    for i in range(week_bron_idx - 1, max(week_bron_idx - 5, -1), -1):
        cell = str(data[i][1]).strip() if len(data[i]) > 1 else ""
        import re
        if re.search(r"\d+\s+неделя", cell, re.IGNORECASE):
            week_label = cell
            break

    # Имена пользователей из строки 12 (index 11)
    header = data[11] if len(data) > 11 else []
    user_names = {}
    for col in USER_COLS_SYNC:
        if col < len(header):
            raw = str(header[col]).strip()
            if raw and raw.lower() not in ("nan", ""):
                clean = re.split(r"\s+(LEADER|JUNIOR|NOT BAD|★|⭑|☆|Группа)", raw)[0].strip()
                user_names[col] = clean

    # Парсим 7 дней
    records = []
    dates   = []
    row_idx = week_bron_idx

    for _ in range(7):
        if row_idx + 3 >= len(data):
            break

        bron_row = data[row_idx]
        plan_row = data[row_idx + 1]
        vol_row  = data[row_idx + 2]
        com_row  = data[row_idx + 3]

        day_ru   = str(plan_row[1]).strip().lower() if len(plan_row) > 1 else ""
        day_en   = DAY_MAP_SYNC.get(day_ru, day_ru.capitalize())
        date_str = str(com_row[1]).strip()  if len(com_row)  > 1 else ""
        time_raw = str(vol_row[1]).strip()  if len(vol_row)  > 1 else ""

        if date_str:
            dates.append(date_str)

        is_no_train  = "нет тренировки" in time_raw.lower()
        is_remote_day = "удал" in time_raw.lower()

        import re as _re
        time_clean = ""
        if not is_no_train and not is_remote_day and time_raw:
            time_clean = _re.sub(r"(\d{1,2})\s+(\d{2})", r"\1:\2", time_raw)

        for col, user in user_names.items():
            booked  = str(bron_row[col]).strip().upper() == "TRUE" if col < len(bron_row) else False
            plan    = str(plan_row[col]).strip() if col < len(plan_row) else ""
            volume  = str(vol_row[col]).strip()  if col < len(vol_row)  else ""
            comment = str(com_row[col]).strip()  if col < len(com_row)  else ""

            plan    = "" if plan    in ("nan", "None") else plan
            volume  = "" if volume  in ("nan", "None") else volume
            comment = "" if comment in ("nan", "None") else comment

            remote = is_remote_day or "удал" in plan.lower() or "удал" in comment.lower()

            records.append([
                user,
                date_str,
                day_en,
                time_clean,
                plan,
                volume,
                "Yes" if remote else "No",
                "TRUE" if booked else "FALSE",
                comment,
            ])

        row_idx += 5

    # Формируем заголовок недели
    week_dates = f"{dates[0]} — {dates[-1]}" if len(dates) >= 2 else ""
    week_line  = f"{week_label}  |  {week_dates}"

    # Пишем в BD
    bd = get_bd_sheet()
    bd.clear()
    bd.update([[week_line]], "A1")
    bd.update([["User", "Date", "Day", "Time", "Plan", "Volume", "Remote", "Booked", "Comments"]], "A2")
    if records:
        bd.update(records, "A3")

# ═══════════════════════════════════════════════════════════════
# БРОНИРОВАНИЕ В ЛИСТЕ 2026
# ═══════════════════════════════════════════════════════════════

def find_date_rows(date_str: str) -> dict | None:
    data = get_source_sheet().get_all_values()
    for i, row in enumerate(data):
        b = str(row[1]).strip() if len(row) > 1 else ""
        if b == date_str:
            bron_1based = i - 2
            vol_1based  = i
            if bron_1based >= 1:
                return {"bron": bron_1based, "vol": vol_1based}
    return None

def set_checkbox(cell_addr: str, value: bool):
    get_source_sheet().update(cell_addr, [[value]], value_input_option="RAW")

def set_booking(user_name: str, date_str: str, value: bool) -> bool:
    rows = find_date_rows(date_str)
    if rows is None:
        return False
    set_checkbox(f"{USER_COLUMNS[user_name]}{rows['bron']}", value)
    return True

def set_remote_booking(user_name: str, date_str: str) -> bool:
    rows = find_date_rows(date_str)
    if rows is None:
        return False
    src = get_source_sheet()
    src.update(f"B{rows['vol']}", [["удаленно"]], value_input_option="RAW")
    set_checkbox(f"{USER_COLUMNS[user_name]}{rows['bron']}", True)
    return True

def cancel_remote_booking(user_name: str, date_str: str) -> bool:
    rows = find_date_rows(date_str)
    if rows is None:
        return False
    src = get_source_sheet()
    time_val = src.acell(f"B{rows['vol']}").value or ""
    if "удал" in time_val.lower():
        src.update(f"B{rows['vol']}", [["нет тренировки"]], value_input_option="RAW")
    set_checkbox(f"{USER_COLUMNS[user_name]}{rows['bron']}", False)
    return True

# ═══════════════════════════════════════════════════════════════
# КЛАВИАТУРЫ
# ═══════════════════════════════════════════════════════════════

def kb_start():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🚀 Начать пользоваться")]],
        resize_keyboard=True,
        one_time_keyboard=True,
        input_field_placeholder="Нажмите кнопку чтобы начать",
    )

def kb_main_menu():
    b = InlineKeyboardBuilder()
    b.button(text="📅 Расписание",  callback_data="schedule")
    b.button(text="📝 Мои записи",  callback_data="my_booking")
    b.button(text="🎫 Абонемент",   callback_data="subscription")
    b.button(text="👤 Профиль",     callback_data="profile")
    b.adjust(2)
    return b.as_markup()

def kb_user_list():
    b = InlineKeyboardBuilder()
    for name in USER_COLUMNS:
        b.button(text=name, callback_data=f"user_{name}")
    b.adjust(2)
    return b.as_markup()

# ═══════════════════════════════════════════════════════════════
# БОТ
# ═══════════════════════════════════════════════════════════════

bot = Bot(token=TOKEN)
dp  = Dispatcher()

# ── /start ──────────────────────────────────────────────────────
@dp.message(CommandStart())
async def cmd_start(message: Message):
    await message.answer_photo(
        photo=FSInputFile("logo.png"),
        caption="Добро пожаловать 👋\nНажмите кнопку ниже чтобы начать.",
        reply_markup=kb_start(),
    )

@dp.message(F.text == "🚀 Начать пользоваться")
async def btn_start(message: Message):
    user_name = get_user_name_by_telegram_id(message.from_user.id)
    if user_name:
        await message.answer_photo(
            photo=FSInputFile("logo.png"),
            caption=f"Привет, {user_name} 👋",
            reply_markup=kb_main_menu(),
        )
        return
    await message.answer_photo(
        photo=FSInputFile("logo.png"),
        caption="Выберите своё имя из списка:",
        reply_markup=kb_user_list(),
    )

# ── Выбор имени ─────────────────────────────────────────────────
@dp.callback_query(F.data.startswith("user_"))
async def cb_select_user(callback: CallbackQuery):
    user_name  = callback.data[5:]
    col_letter = USER_COLUMNS[user_name]
    current    = get_source_sheet().acell(f"{col_letter}6").value
    my_id      = str(callback.from_user.id)

    if str(current).strip() == my_id:
        await callback.message.edit_caption(caption=f"Привет, {user_name} 👋", reply_markup=kb_main_menu())
        await callback.answer()
        return

    warning = ""
    if current and current.strip() not in ("", "None", "0"):
        warning = "\n\n⚠️ Это имя уже привязано к другому аккаунту."

    b = InlineKeyboardBuilder()
    b.button(text="✅ Подтвердить", callback_data=f"confirm_{user_name}_{my_id}")
    b.button(text="❌ Отмена",      callback_data="back_to_list")
    b.adjust(2)

    await callback.message.edit_caption(caption=
        f"Вы выбрали: *{user_name}*{warning}\nПодтвердить регистрацию?",
        parse_mode="Markdown",
        reply_markup=b.as_markup(),
    )
    await callback.answer()

@dp.callback_query(F.data.startswith("confirm_"))
async def cb_confirm(callback: CallbackQuery):
    parts     = callback.data.split("_")
    user_name = parts[1]
    user_id   = int(parts[2])
    save_telegram_id(user_name, user_id)
    await callback.message.edit_caption(caption=
        f"✅ Готово! Добро пожаловать, {user_name} 👋",
        reply_markup=kb_main_menu(),
    )
    await callback.answer()

@dp.callback_query(F.data == "back_to_list")
async def cb_back_to_list(callback: CallbackQuery):
    await callback.message.edit_caption(caption="Выберите своё имя:", reply_markup=kb_user_list())
    await callback.answer()

# ── Шаг 1: компактный список дней ───────────────────────────────
@dp.callback_query(F.data == "schedule")
async def cb_schedule(callback: CallbackQuery, _toast: str = ""):
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    _ensure_bd()

    if not _bd_rows or not _bd_rows[0]:
        b = InlineKeyboardBuilder()
        b.button(text="◀️ Назад", callback_data="main_menu")
        await callback.message.edit_caption(
            caption="⏳ Расписание ещё не готово.\nТренер скоро откроет новую неделю.",
            reply_markup=b.as_markup(),
        )
        await callback.answer()
        return

    week   = get_week_info()
    trains = get_schedule_for_user(user_name)
    booked = [t for t in trains if t["booked"]]

    text = f"🗓 {week['label']}   {week['dates']}\n\n"
    if booked:
        booked_labels = []
        for t in booked:
            short = DAY_SHORT.get(t["day"], t["day"])
            booked_labels.append(f"{short} {t['date'][:5]}")
        text += f"✅ Вы записаны: {', '.join(booked_labels)}\n\n"
    text += "Нажмите на день чтобы записаться или отменить:\n"

    b = InlineKeyboardBuilder()
    for t in trains:
        short      = DAY_SHORT.get(t["day"], t["day"])
        date_short = t["date"][:5]

        if t["booked"]:
            if t["remote"] or not t["time"]:
                label = f"✅  {short} {date_short} — 🏠 удалённо"
            else:
                label = f"✅  {short} {date_short} — 🏊 {t['time']}"
        elif t["time"]:
            label = f"🏊  {short} {date_short} — {t['time']}"
        else:
            label = f"🏠  {short} {date_short} — можно удалённо"

        b.button(text=label, callback_data=f"day_{t['date']}")

    b.button(text="🔄 Обновить", callback_data="refresh_schedule")
    b.button(text="◀️ Назад", callback_data="main_menu")
    b.adjust(1)

    try:
        await callback.message.edit_caption(caption=text, reply_markup=b.as_markup())
    except TelegramBadRequest:
        pass  # содержимое не изменилось — игнорируем
    await callback.answer(_toast)

# ── Обновить расписание ──────────────────────────────────────────
@dp.callback_query(F.data == "refresh_schedule")
async def cb_refresh_schedule(callback: CallbackQuery):
    _invalidate_bd()
    await cb_schedule(callback, "✅ Расписание обновлено")

# ── Шаг 2: карточка дня + подтверждение ─────────────────────────
@dp.callback_query(F.data.startswith("day_"))
async def cb_day_detail(callback: CallbackQuery):
    date_str  = callback.data[4:]
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    _ensure_bd()
    trains = get_schedule_for_user(user_name)
    t = next((x for x in trains if x["date"] == date_str), None)
    if not t:
        await callback.message.edit_caption(caption="❌ День не найден.")
        await callback.answer()
        return

    is_remote_day = not t["time"]

    if t["booked"]:
        if is_remote_day or t["remote"]:
            header = f"✅ {t['day']} {date_str} — 🏠 удалённо"
        else:
            header = f"✅ {t['day']} {date_str} — 🏊 {t['time']}"
    elif is_remote_day:
        header = f"🏠 {t['day']} {date_str} — можно записаться удалённо"
    else:
        header = f"🏊 {t['day']} {date_str} — {t['time']}"

    text = f"*{header}*\n\n"

    if t["plan"]:
        text += f"📋 *План:*\n{t['plan']}\n\n"
    elif t["booked"]:
        text += "📋 *План:* тренер ещё не написал — ожидайте\n\n"

    if t["volume"]:
        text += f"📏 *Объём:* {t['volume']}\n"

    if t["comments"]:
        text += f"💬 *Комментарий:* {t['comments']}\n"

    b = InlineKeyboardBuilder()
    if t["booked"]:
        b.button(text="❌ Отменить запись", callback_data=f"unbook_confirm_{date_str}")
    elif is_remote_day:
        b.button(text="🏠 Записаться удалённо", callback_data=f"book_remote_{date_str}")
    else:
        b.button(text="🏊 Записаться", callback_data=f"book_{date_str}")
    b.button(text="◀️ К расписанию", callback_data="schedule")
    b.adjust(1)

    await callback.message.edit_caption(caption=text, parse_mode="Markdown", reply_markup=b.as_markup())
    await callback.answer()

# ── Бронирование удалённо ────────────────────────────────────────
@dp.callback_query(F.data.startswith("book_remote_"))
async def cb_book_remote(callback: CallbackQuery):
    date_str  = callback.data[12:]
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    ok = set_remote_booking(user_name, date_str)
    if ok:
        _invalidate_bd()
        await callback.message.edit_caption(caption=
            f"🏠 Вы записались на удалённую тренировку {date_str}!\n"
            f"Тренер напишет план перед тренировкой.",
            reply_markup=kb_main_menu(),
        )
        _last_booking[user_name] = date_str
        _inactivity_notified.discard(user_name)
        save_last_booking(user_name, date_str)
        await notify_trainer(f"🏠 {user_name} записался удалённо — {date_str}")
        await check_subscription(user_name)
    else:
        await callback.message.edit_caption(caption="❌ Не удалось записаться. Обратитесь к тренеру.")
    await callback.answer()

# ── Бронирование в бассейн ───────────────────────────────────────
@dp.callback_query(F.data.startswith("book_") & ~F.data.startswith("book_remote_"))
async def cb_book(callback: CallbackQuery):
    date_str  = callback.data[5:]
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    ok = set_booking(user_name, date_str, True)
    if ok:
        _invalidate_bd()
        await callback.message.edit_caption(caption=
            f"🏊 Вы записаны на тренировку {date_str}!",
            reply_markup=kb_main_menu(),
        )
        _last_booking[user_name] = date_str
        _inactivity_notified.discard(user_name)
        save_last_booking(user_name, date_str)
        await notify_trainer(f"🏊 {user_name} забронировал тренировку — {date_str}")
        await check_subscription(user_name)
    else:
        await callback.message.edit_caption(caption="❌ Не удалось записаться. Обратитесь к тренеру.")
    await callback.answer()

# ── Мои записи ──────────────────────────────────────────────────
@dp.callback_query(F.data == "my_booking")
async def cb_my_booking(callback: CallbackQuery):
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    _ensure_bd()

    if not _bd_rows or not _bd_rows[0]:
        b = InlineKeyboardBuilder()
        b.button(text="◀️ Назад", callback_data="main_menu")
        await callback.message.edit_caption(
            caption="⏳ Расписание ещё не готово.\nТренер скоро откроет новую неделю.",
            reply_markup=b.as_markup(),
        )
        await callback.answer()
        return

    week   = get_week_info()
    trains = get_schedule_for_user(user_name)
    booked = [t for t in trains if t["booked"]]

    if not booked:
        b = InlineKeyboardBuilder()
        b.button(text="📅 Перейти к расписанию", callback_data="schedule")
        b.button(text="◀️ Назад", callback_data="main_menu")
        b.adjust(1)
        await callback.message.edit_caption(caption=
            f"📝 *{week['label']}* — у вас пока нет записей.\n\n"
            f"Запишитесь через раздел 📅 Расписание.",
            parse_mode="Markdown",
            reply_markup=b.as_markup(),
        )
        await callback.answer()
        return

    text = f"📝 *{week['label']}*   _{week['dates']}_\n\n"
    for t in booked:
        time_display = t["time"] if t["time"] else "удалённо"
        fmt = " 🏠" if (t["remote"] or not t["time"]) else " 🏊"
        text += f"✅ *{t['day']}* {t['date']} — {time_display}{fmt}\n"

        if t["plan"]:
            text += f"    📋 _План:_\n{t['plan']}\n"
        else:
            text += "    📋 _План:_ тренер ещё не написал\n"

        if t["volume"]:
            text += f"    📏 _Объём:_ {t['volume']}\n"
        else:
            text += "    📏 _Объём:_ —\n"

        if t["comments"]:
            text += f"    💬 _{t['comments']}_\n"

        text += "\n"

    text += "Нажмите на тренировку чтобы отменить запись."

    b = InlineKeyboardBuilder()
    for t in booked:
        time_display = t["time"] if t["time"] else "удалённо"
        short = DAY_SHORT.get(t["day"], t["day"])
        b.button(
            text=f"❌ Отменить — {short} {t['date'][:5]}",
            callback_data=f"unbook_confirm_{t['date']}"
        )
    b.button(text="◀️ Назад", callback_data="main_menu")
    b.adjust(1)

    await callback.message.edit_caption(caption=text, parse_mode="Markdown", reply_markup=b.as_markup())
    await callback.answer()

# ── Подтверждение отмены ─────────────────────────────────────────
@dp.callback_query(F.data.startswith("unbook_confirm_"))
async def cb_unbook_confirm(callback: CallbackQuery):
    date_str  = callback.data[15:]
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    _ensure_bd()
    trains = get_schedule_for_user(user_name)
    t = next((x for x in trains if x["date"] == date_str), None)
    time_display = t["time"] if t and t["time"] else "удалённо"
    day_name = t["day"] if t else date_str

    b = InlineKeyboardBuilder()
    b.button(text="✅ Да, отменить",    callback_data=f"unbook_{date_str}")
    b.button(text="◀️ Нет, оставить",  callback_data="my_booking")
    b.adjust(1)

    await callback.message.edit_caption(caption=
        f"Вы уверены что хотите отменить запись?\n\n"
        f"*{day_name} {date_str} — {time_display}*",
        parse_mode="Markdown",
        reply_markup=b.as_markup(),
    )
    await callback.answer()

# ── Подтверждённая отмена ────────────────────────────────────────
@dp.callback_query(F.data.startswith("unbook_") & ~F.data.startswith("unbook_confirm_"))
async def cb_unbook(callback: CallbackQuery):
    date_str  = callback.data[7:]
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    rows = find_date_rows(date_str)
    is_remote = False
    if rows:
        time_val = get_source_sheet().acell(f"B{rows['vol']}").value or ""
        is_remote = "удал" in time_val.lower()

    ok = cancel_remote_booking(user_name, date_str) if is_remote else set_booking(user_name, date_str, False)

    if ok:
        _invalidate_bd()
        await callback.message.edit_caption(caption=
            f"✅ Запись на {date_str} отменена.",
            reply_markup=kb_main_menu(),
        )
        await notify_trainer(f"❌ {user_name} отменил запись — {date_str}")
    else:
        await callback.message.edit_caption(caption="❌ Не удалось отменить. Обратитесь к тренеру.")
    await callback.answer()

# ── Абонемент ────────────────────────────────────────────────────
@dp.callback_query(F.data == "subscription")
async def cb_subscription(callback: CallbackQuery):
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    col  = USER_COLUMNS[user_name]
    src  = get_source_sheet()
    # Строка 9 — остаток абонемента
    data      = src.batch_get([f"{col}9"])
    remaining = data[0][0][0] if data[0] else "—"

    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data="main_menu")

    await callback.message.edit_caption(caption=
        f"🎫 *Абонемент — {user_name}*\n\n"
        f"Остаток: *{remaining}* тренировок",
        parse_mode="Markdown",
        reply_markup=b.as_markup(),
    )
    await callback.answer()

# ── Профиль ─────────────────────────────────────────────────────
@dp.callback_query(F.data == "profile")
async def cb_profile(callback: CallbackQuery):
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    col  = USER_COLUMNS[user_name]
    src  = get_source_sheet()
    # Строка 4 — метры за 2025, строка 5 — метры за 2026
    # Метры за неделю считаем из листа BD (сумма Volume для записанных тренировок)
    data = src.batch_get([f"{col}4", f"{col}5"])
    meters_last_year = data[0][0][0] if data[0] else "—"
    meters_year      = data[1][0][0] if data[1] else "—"

    # Метры за текущую неделю — сумма из BD по записанным тренировкам
    try:
        trains = get_schedule_for_user(user_name)
        week_meters = sum(
            int(str(t["volume"]).replace(" м", "").replace(",", "").strip())
            for t in trains if t["booked"] and t["volume"]
        )
        meters_week = f"{week_meters} м" if week_meters > 0 else "—"
    except Exception:
        meters_week = "—"

    # Сравнение с прошлым годом
    try:
        this_y = int(str(meters_year).replace(" ", "").replace(",", ""))
        last_y = int(str(meters_last_year).replace(" ", "").replace(",", ""))
        diff   = this_y - last_y
        arrow  = "📈" if diff >= 0 else "📉"
        compare = f"{arrow} {'+' if diff >= 0 else ''}{diff:,} м к прошлому году".replace(",", " ")
    except Exception:
        compare = ""

    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data="main_menu")

    await callback.message.edit_caption(caption=
        f"👤 *Профиль — {user_name}*\n\n"
        f"🏊 Метров на этой неделе:\n*{meters_week}*\n\n"
        f"📅 Метров в {2026} году:\n*{meters_year}*\n\n"
        f"📅 Метров в {2025} году:\n*{meters_last_year}*\n"
        f"{compare}",
        parse_mode="Markdown",
        reply_markup=b.as_markup(),
    )
    await callback.answer()

# ── Главное меню ─────────────────────────────────────────────────
@dp.callback_query(F.data == "main_menu")
async def cb_main_menu(callback: CallbackQuery):
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    name = user_name or "пользователь"
    await callback.message.edit_caption(caption=f"Главное меню, {name} 👋", reply_markup=kb_main_menu())
    await callback.answer()

# ═══════════════════════════════════════════════════════════════
# ЗАПУСК
# ═══════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════
# УВЕДОМЛЕНИЯ
# ═══════════════════════════════════════════════════════════════

async def notify_trainer(text: str):
    try:
        await bot.send_message(TRAINER_ID, text)
    except Exception as e:
        print(f"notify_trainer ошибка: {e}")

async def notify_user(user_name: str, text: str, parse_mode: str | None = None):
    tid = _tid_cache.get(user_name)
    if not tid:
        return
    try:
        await bot.send_message(tid, text, parse_mode=parse_mode)
    except Exception as e:
        print(f"notify_user({user_name}) ошибка: {e}")

async def check_subscription(user_name: str):
    try:
        col       = USER_COLUMNS[user_name]
        remaining = get_source_sheet().acell(f"{col}9").value or "0"
        left      = int(str(remaining).replace(" ", "").replace(",", "") or "0")
        if left < 0:
            await notify_user(user_name,
                f"🚨 {user_name}, вы ушли в минус: {left} тренировок.\n"
                f"Свяжитесь с тренером для пополнения абонемента.")
        elif left == 0:
            await notify_user(user_name,
                f"🔴 {user_name}, абонемент закончился.\n"
                f"Свяжитесь с тренером для пополнения.")
        elif left <= 2:
            await notify_user(user_name,
                f"🟡 {user_name}, осталось {left} тренировки — скоро закончится абонемент.")
    except Exception as e:
        print(f"check_subscription({user_name}) ошибка: {e}")

async def weekly_report():
    sent_this_week = False
    while True:
        await asyncio.sleep(60)
        try:
            now = datetime.now()
            # Воскресенье (weekday=6) в 20:00
            if now.weekday() == 6 and now.hour == 20 and now.minute == 0:
                if sent_this_week:
                    continue
                sent_this_week = True
                _ensure_bd()
                for user_name in USER_COLUMNS:
                    trains = get_schedule_for_user(user_name)
                    booked = [t for t in trains if t["booked"]]
                    if not booked:
                        continue
                    total_meters = 0
                    for t in booked:
                        try:
                            total_meters += int(
                                str(t["volume"]).replace(" м", "").replace(",", "").strip()
                            )
                        except Exception:
                            pass
                    if total_meters == 0:
                        continue
                    week = get_week_info()
                    text = (
                        f"📊 Итог недели, {user_name}!\n\n"
                        f"🏊 Тренировок: {len(booked)}\n"
                        f"📏 Метров: {total_meters:,} м\n\n"
                        f"Неделя: {week['label']} {week['dates']}"
                    ).replace(",", " ")
                    await notify_user(user_name, text)
            else:
                # Сбрасываем флаг в понедельник
                if now.weekday() == 0 and now.hour == 0:
                    sent_this_week = False
        except Exception as e:
            print(f"weekly_report ошибка: {e}")

async def inactivity_checker():
    """Ежедневно в 10:00 проверяет пользователей которые не брали тренировки 14+ дней."""
    while True:
        await asyncio.sleep(60)
        try:
            now = datetime.now()
            if now.hour != 10 or now.minute != 0:
                continue
            for user_name, last_date_str in list(_last_booking.items()):
                if user_name in _inactivity_notified:
                    continue
                try:
                    last_date = datetime.strptime(last_date_str, "%d.%m.%Y")
                except ValueError:
                    try:
                        last_date = datetime.strptime(last_date_str + f".{now.year}", "%d.%m.%Y")
                    except Exception:
                        continue
                days_gone = (now - last_date).days
                if days_gone >= 14:
                    _inactivity_notified.add(user_name)
                    await notify_user(
                        user_name,
                        f"👋 Привет, {user_name}!\n\n"
                        f"Давно не виделись — уже {days_gone} дней без тренировки.\n"
                        f"Может поплаваем? 🏊"
                    )
        except Exception as e:
            print(f"inactivity_checker ошибка: {e}")

async def training_reminder():
    """За 2 часа до тренировки отправляет напоминание записавшимся пользователям."""
    global _reminded_training, _reminder_date
    while True:
        await asyncio.sleep(60)
        try:
            now = datetime.now()
            today_str = now.strftime("%d.%m.%Y")

            if today_str != _reminder_date:
                _reminded_training = set()
                _reminder_date = today_str

            _ensure_bd()
            if not _bd_rows or not _bd_rows[0]:
                continue

            for user_name in USER_COLUMNS:
                trains = get_schedule_for_user(user_name)
                for t in trains:
                    if not t["booked"] or not t["time"] or t["date"] != today_str:
                        continue
                    key = f"{user_name}|{t['date']}"
                    if key in _reminded_training:
                        continue
                    try:
                        h, m = map(int, t["time"].split(":"))
                        train_dt = now.replace(hour=h, minute=m, second=0, microsecond=0)
                        delta_min = (train_dt - now).total_seconds() / 60
                        if 110 <= delta_min <= 130:
                            _reminded_training.add(key)
                            await notify_user(
                                user_name,
                                f"⏰ {user_name}, напоминание!\n\n"
                                f"Через ~2 часа тренировка в {t['time']} 🏊\n"
                                f"{t['day']}, {t['date']}"
                            )
                    except Exception:
                        continue
        except Exception as e:
            print(f"training_reminder ошибка: {e}")

async def plan_checker():
    """Каждые 5 минут проверяет появился ли новый план для записанных пользователей."""
    # Снимаем baseline сразу при старте (без задержки)
    try:
        _ensure_bd()
        if _bd_rows and _bd_rows[0]:
            for user_name in USER_COLUMNS:
                for t in get_schedule_for_user(user_name):
                    if t["booked"]:
                        key = f"{user_name}|{t['date']}"
                        _known_plans[key] = t["plan"].strip()
    except Exception as e:
        print(f"plan_checker (baseline) ошибка: {e}")

    while True:
        await asyncio.sleep(300)
        try:
            _invalidate_bd()
            _ensure_bd()
            if not _bd_rows or not _bd_rows[0]:
                continue

            for user_name in USER_COLUMNS:
                trains = get_schedule_for_user(user_name)
                for t in trains:
                    if not t["booked"]:
                        continue
                    key = f"{user_name}|{t['date']}"
                    new_plan = t["plan"].strip()
                    old_plan = _known_plans.get(key, "")
                    if new_plan and new_plan != old_plan:
                        _known_plans[key] = new_plan
                        time_label = t["time"] if t["time"] else "удалённо"
                        await notify_user(
                            user_name,
                            f"📋 {user_name}, тренер написал план!\n\n"
                            f"{t['day']} {t['date']} — {time_label}\n\n"
                            f"{new_plan}",
                        )
        except Exception as e:
            print(f"plan_checker ошибка: {e}")


async def week_watcher():
    global _week_marker_row, _week_session_notified
    while True:
        await asyncio.sleep(15)
        try:
            data = get_source_sheet().get_all_values()

            new_row   = -1
            new_label = ""
            for i, row in enumerate(data):
                last = str(row[-1]).strip().lower() if row else ""
                if "текущ" in last:
                    new_row = i
                    import re
                    for j in range(i - 1, max(i - 5, -1), -1):
                        cell = str(data[j][1]).strip() if len(data[j]) > 1 else ""
                        if re.search(r"\d+\s+неделя", cell, re.IGNORECASE):
                            new_label = cell
                            break
                    break

            is_active = new_row != -1

            # Сбрасываем кэш если строка изменилась
            if new_row != _week_marker_row:
                print(f"Смена недели: строка {_week_marker_row} → {new_row}, сбрасываю кэш")
                _week_marker_row = new_row
                _invalidate_bd()

            # Когда неделя стала неактивной — сбрасываем флаг уведомления
            if not is_active:
                _week_session_notified = False

            # Когда неделя активна и ещё не уведомляли — шлём
            elif is_active and not _week_session_notified:
                _week_session_notified = True
                print(f"Расписание открыто: «{new_label}», уведомляю пользователей")
                load_all_tids()
                _ensure_bd()
                for user_name in USER_COLUMNS:
                    await notify_user(
                        user_name,
                        f"📅 {user_name}, расписание на {new_label} готово!\n"
                        f"Открой бот чтобы посмотреть и записаться 👇"
                    )
        except Exception as e:
            print(f"week_watcher ошибка: {e}")

async def main():
    print("Бот запущен 🚀")
    load_all_tids()
    load_last_bookings()
    asyncio.create_task(week_watcher())
    asyncio.create_task(weekly_report())
    asyncio.create_task(inactivity_checker())
    asyncio.create_task(training_reminder())
    asyncio.create_task(plan_checker())
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())