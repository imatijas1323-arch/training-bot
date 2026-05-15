"""
bot.py — Swimming Training Bot
"""

import asyncio
import os
import re
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, CallbackQuery, ReplyKeyboardMarkup, KeyboardButton, FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton
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
BOT_TIMEZONE   = os.getenv("BOT_TIMEZONE", "Europe/Moscow")
BOT_TZ         = ZoneInfo(BOT_TIMEZONE)

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

SWIM_GRADE_ORDER = [
    "JUNIOR",
    "NOT BAD ⭑⭑⭑",
    "NOT BAD ★★",
    "NOT BAD ☆",
    "LEADER ⭑⭑⭑",
    "LEADER ★★",
    "LEADER ☆",
    "ELITE",
    "PRO",
]

SWIM_GRADE_INFO = {
    "JUNIOR": (
        "🏊 *JUNIOR*\n\n"
        "Уверенно плывёт 50м вольным стилем.\n"
        "Правильное дыхание, базовая техника."
    ),
    "NOT BAD ⭑⭑⭑": (
        "🏊 *NOT BAD ⭑⭑⭑*\n\n"
        "📏 100м вольным стилем\n"
        "⏱ Темп 50м: 55 с (М) / 60 с (Ж)"
    ),
    "NOT BAD ★★": (
        "🏊 *NOT BAD ★★*\n\n"
        "📏 200м вольным стилем\n"
        "⏱ Темп 50м: 52 с (М) / 58 с (Ж)"
    ),
    "NOT BAD ☆": (
        "🏊 *NOT BAD ☆*\n\n"
        "📏 400м вольным стилем\n"
        "⏱ Темп 50м: 49 с (М) / 52 с (Ж)"
    ),
    "LEADER ⭑⭑⭑": (
        "🏊 *LEADER ⭑⭑⭑*\n\n"
        "📏 1000м вольным стилем\n"
        "⏱ Темп 50м: 45 с (М) / 49 с (Ж)"
    ),
    "LEADER ★★": (
        "🏊 *LEADER ★★*\n\n"
        "📏 1000м до 23 мин (М) / 25 мин (Ж)\n"
        "⏱ Темп 50м: 41 с (М) / 45 с (Ж)"
    ),
    "LEADER ☆": (
        "🏊 *LEADER ☆*\n\n"
        "📏 2000м вольным стилем\n"
        "⏱ Темп 50м: 36 с (М)"
    ),
    "ELITE": (
        "🏊 *ELITE*\n\n"
        "📏 4000м вольным стилем\n"
        "⏱ Темп 50м: 29 с (М) / 33 с (Ж)\n"
        "⏱ 1000м до 16 мин (М) / 19 мин (Ж)"
    ),
    "PRO": (
        "🏊 *PRO*\n\n"
        "Высший уровень — соревновательные ранги и титулы."
    ),
}

DNF_GRADE_ORDER = ["★", "★★", "★★★"]

DNF_GRADE_INFO = {
    "★": (
        "🤿 *DNF/DYN ★*\n\n"
        "🫁 DNF (без ласт): 25 м\n"
        "🦈 DYN (с ластами): 35 м"
    ),
    "★★": (
        "🤿 *DNF/DYN ★★*\n\n"
        "🫁 DNF (без ласт): 35 м\n"
        "🦈 DYN (с ластами): 50 м"
    ),
    "★★★": (
        "🤿 *DNF/DYN ★★★*\n\n"
        "🫁 DNF (без ласт): 50 м\n"
        "🦈 DYN (с ластами): 75 м"
    ),
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

def get_meta_sheet():
    for ws in ss.worksheets():
        if ws.title == "Meta":
            return ws
    ws = ss.add_worksheet(title="Meta", rows=100, cols=2)
    ws.update([["user", "last_booking"]], "A1")
    return ws

def get_grades_sheet():
    for ws in ss.worksheets():
        if ws.title == "Grades":
            return ws
    ws = ss.add_worksheet(title="Grades", rows=1000, cols=3)
    ws.update([["user", "grade", "date"]], "A1")
    return ws

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
_state_notified:      set[str]       = set() # "имя|дата" — кому уже отправили запрос о состоянии
_state_date:          str            = ""    # дата для сброса _state_notified в полночь
_multi_select:        dict[int, dict[str, str]] = {}  # telegram_id → {date: "pool"|"remote"|"pending"}
_known_grades:        dict[str, str]            = {}  # имя → текущий грейд плавания
_known_dnf_grades:    dict[str, str]            = {}  # имя → текущий DNF/DYN грейд
BD_TTL = 300                                 # секунд (5 минут)

STATES = [
    "-",
    "1-4 восстановление/очень легко",
    "5 легко",
    "6 умеренно",
    "7 хорошая работа",
    "8 довольно тяжело",
    "9 тяжело",
    "10 максимальное усилие",
]

def _now() -> datetime:
    return datetime.now(BOT_TZ)

def _parse_sheet_date(date_str: str, year: int):
    date_str = str(date_str).strip()
    for fmt, value in (
        ("%d.%m.%Y", date_str),
        ("%d.%m.%Y", f"{date_str}.{year}"),
    ):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None

def _is_today(date_str: str, now: datetime) -> bool:
    sheet_date = _parse_sheet_date(date_str, now.year)
    return sheet_date == now.date()

def _is_past_day(date_str: str) -> bool:
    """True если дата строго раньше сегодня."""
    d = _parse_sheet_date(date_str, _now().year)
    return d is not None and d < _now().date()

def _is_finished(date_str: str, time_str: str = "") -> bool:
    """True если тренировка закончилась: прошедший день, или сегодня и время+2ч истекло."""
    from datetime import timedelta
    now = _now()
    d = _parse_sheet_date(date_str, now.year)
    if d is None:
        return False
    if d < now.date():
        return True
    if d > now.date():
        return False
    if not time_str:
        return False
    try:
        h, m = map(int, time_str.split(":"))
        train_dt = now.replace(hour=h, minute=m, second=0, microsecond=0)
        return now >= train_dt + timedelta(hours=2)
    except Exception:
        return False

def _invalidate_bd():
    global _bd_ts
    _bd_ts = 0.0

def load_all_tids():
    """Загружает telegram_id всех пользователей из строки 6 таблицы в кэш."""
    try:
        row = get_source_sheet().row_values(6)
        count = 0
        for name, col_letter in USER_COLUMNS.items():
            col_idx = ord(col_letter) - ord("A")
            if col_idx < len(row):
                val = str(row[col_idx]).strip()
                if val and val not in ("", "0", "None"):
                    try:
                        tid = int(val)
                        _tid_cache[name]  = tid
                        _user_cache[tid]  = name
                        count += 1
                    except ValueError:
                        pass
        print(f"load_all_tids: загружено {count} пользователей")
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

def _extract_swim_grade(text: str) -> str:
    """Извлекает грейд плавания из текста (левая часть до '/')."""
    for keyword, symbol in [
        ("NOT BAD", "⭑⭑⭑"), ("NOT BAD", "★★"), ("NOT BAD", "☆"),
        ("LEADER",  "⭑⭑⭑"), ("LEADER",  "★★"), ("LEADER",  "☆"),
    ]:
        if keyword in text and symbol in text:
            return f"{keyword} {symbol}"
    for grade in ("PRO", "ELITE", "LEADER", "NOT BAD", "JUNIOR"):
        if grade in text:
            return grade
    return ""

def _extract_dnf_grade(text: str) -> str:
    """Извлекает DNF/DYN грейд из текста (правая часть после '/')."""
    text = text.strip()
    if "★★★" in text: return "★★★"
    if "★★"  in text: return "★★"
    if "★"   in text: return "★"
    return ""

def _parse_grade(raw: str) -> str:
    """Грейд плавания из ячейки (до '/')."""
    left = raw.split("/")[0] if "/" in raw else raw
    return _extract_swim_grade(left)

def _parse_dnf_grade(raw: str) -> str:
    """DNF/DYN грейд из ячейки (после '/')."""
    if "/" not in raw:
        return ""
    return _extract_dnf_grade(raw.split("/", 1)[1])

def get_user_grade(user_name: str) -> str:
    if user_name in _known_grades:
        return _known_grades[user_name]
    try:
        col = USER_COLUMNS[user_name]
        raw = str(get_source_sheet().acell(f"{col}12").value or "").strip()
        return _parse_grade(raw)
    except Exception:
        return ""

def get_user_dnf_grade(user_name: str) -> str:
    if user_name in _known_dnf_grades:
        return _known_dnf_grades[user_name]
    try:
        col = USER_COLUMNS[user_name]
        raw = str(get_source_sheet().acell(f"{col}12").value or "").strip()
        return _parse_dnf_grade(raw)
    except Exception:
        return ""

def load_grades():
    """Загружает грейды из строки 12 в кэш и пишет в Grades тех у кого нет записи."""
    try:
        row = get_source_sheet().row_values(12)
        grades_sheet = get_grades_sheet()
        existing = grades_sheet.get_all_values()
        users_with_entry = {r[0] for r in existing[1:] if r}
        date_str = _now().strftime("%d.%m.%Y")
        rows_to_add = []
        for name, col_letter in USER_COLUMNS.items():
            col_idx = ord(col_letter) - ord("A")
            raw = str(row[col_idx]).strip() if col_idx < len(row) else ""
            grade = _parse_grade(raw)
            _known_grades[name] = grade
            _known_dnf_grades[name] = _parse_dnf_grade(raw)
            if grade and name not in users_with_entry:
                rows_to_add.append([name, grade, date_str])
        if rows_to_add:
            grades_sheet.append_rows(rows_to_add)
    except Exception as e:
        print(f"load_grades ошибка: {e}")

def save_grade_history(user_name: str, grade: str):
    try:
        date_str = _now().strftime("%d.%m.%Y")
        get_grades_sheet().append_row([user_name, grade, date_str])
    except Exception as e:
        print(f"save_grade_history({user_name}) ошибка: {e}")

def save_last_booking(user_name: str, date_str: str):
    try:
        meta = get_meta_sheet()
        rows = meta.get_all_values()
        for i, row in enumerate(rows):
            if row and row[0] == user_name:
                meta.update(f"B{i + 1}", [[date_str]])
                return
        meta.append_row([user_name, date_str])
    except Exception as e:
        print(f"save_last_booking({user_name}) ошибка: {e}")

def load_last_bookings():
    """Загружает даты последних бронирований из листа Meta."""
    try:
        rows = get_meta_sheet().get_all_values()
        for row in rows:
            if len(row) >= 2 and row[0] in USER_COLUMNS:
                val = str(row[1]).strip()
                if val:
                    _last_booking[row[0]] = val
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
                clean = re.split(r"\s+(PRO|ELITE|LEADER|JUNIOR|NOT BAD|★|⭑|☆|Группа)", raw)[0].strip()
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

        time_clean = ""
        if not is_no_train and not is_remote_day and time_raw:
            time_clean = re.sub(r"(\d{1,2})\s+(\d{2})", r"\1:\2", time_raw)

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

def set_state(user_name: str, date_str: str, state_value: str) -> bool:
    rows = find_date_rows(date_str)
    if rows is None:
        return False
    try:
        state_row = rows["bron"] + 4
        get_source_sheet().update(f"{USER_COLUMNS[user_name]}{state_row}", [[state_value]])
        return True
    except Exception as e:
        print(f"set_state({user_name}, {date_str}) ошибка: {e}")
        return False

def set_booking(user_name: str, date_str: str, value: bool) -> bool:
    rows = find_date_rows(date_str)
    if rows is None:
        return False
    set_checkbox(f"{USER_COLUMNS[user_name]}{rows['bron']}", value)
    return True

def set_remote_booking_comment(user_name: str, date_str: str) -> bool:
    """Бронирует удалённо: пишет 'удалённо' в ячейку комментария пользователя, не трогает колонку B."""
    rows = find_date_rows(date_str)
    if rows is None:
        return False
    try:
        comment_row = rows["vol"] + 1
        get_source_sheet().update(f"{USER_COLUMNS[user_name]}{comment_row}", [["удалённо"]])
        set_checkbox(f"{USER_COLUMNS[user_name]}{rows['bron']}", True)
        return True
    except Exception as e:
        print(f"set_remote_booking_comment({user_name}, {date_str}) ошибка: {e}")
        return False

def cancel_booking(user_name: str, date_str: str) -> bool:
    """Снимает чекбокс бронирования и очищает комментарий пользователя."""
    rows = find_date_rows(date_str)
    if rows is None:
        return False
    try:
        comment_row = rows["vol"] + 1
        src = get_source_sheet()
        src.update(f"{USER_COLUMNS[user_name]}{comment_row}", [[""]])
        set_checkbox(f"{USER_COLUMNS[user_name]}{rows['bron']}", False)
        return True
    except Exception as e:
        print(f"cancel_booking({user_name}, {date_str}) ошибка: {e}")
        return False

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

def kb_persistent():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🏠 Главное меню")]],
        resize_keyboard=True,
    )

def kb_main_menu():
    b = InlineKeyboardBuilder()
    b.button(text="📅 Расписание",  callback_data="schedule")
    b.button(text="📝 Мои записи",  callback_data="my_booking")
    b.button(text="🎫 Абонемент",   callback_data="subscription")
    b.button(text="👤 Профиль",     callback_data="profile")
    b.button(text="🎯 Мой уровень", callback_data="my_level")
    b.adjust(2, 2, 1)
    return b.as_markup()

def kb_user_list():
    b = InlineKeyboardBuilder()
    for name in USER_COLUMNS:
        b.button(text=name, callback_data=f"user_{name}")
    b.button(text="❓ Меня нет в списке", callback_data="not_in_list")
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
    user_name = get_user_name_by_telegram_id(message.from_user.id)
    if user_name:
        await message.answer("—", reply_markup=kb_persistent())
        await message.answer_photo(
            photo=FSInputFile("logo.png"),
            caption=f"Привет, {user_name} 👋",
            reply_markup=kb_main_menu(),
        )
    else:
        await message.answer_photo(
            photo=FSInputFile("logo.png"),
            caption="Добро пожаловать 👋\nНажмите кнопку ниже чтобы начать.",
            reply_markup=kb_start(),
        )

@dp.message(F.text == "🚀 Начать пользоваться")
async def btn_start(message: Message):
    user_name = get_user_name_by_telegram_id(message.from_user.id)
    if user_name:
        await message.answer("—", reply_markup=kb_persistent())
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

@dp.message(F.text == "🏠 Главное меню")
async def btn_main_menu(message: Message):
    user_name = get_user_name_by_telegram_id(message.from_user.id)
    if user_name:
        await message.answer_photo(
            photo=FSInputFile("logo.png"),
            caption=f"Привет, {user_name} 👋",
            reply_markup=kb_main_menu(),
        )
    else:
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
    await callback.message.answer("👇", reply_markup=kb_persistent())
    tg_user = callback.from_user
    username_str = f"@{tg_user.username}" if tg_user.username else "нет username"
    await notify_trainer(
        f"🆕 Новый пользователь зарегистрировался!\n"
        f"Имя: {user_name}\n"
        f"TG: {tg_user.full_name} ({username_str})\n"
        f"ID: {user_id}"
    )
    await callback.answer()

@dp.callback_query(F.data == "back_to_list")
async def cb_back_to_list(callback: CallbackQuery):
    await callback.message.edit_caption(caption="Выберите своё имя:", reply_markup=kb_user_list())
    await callback.answer()

@dp.callback_query(F.data == "not_in_list")
async def cb_not_in_list(callback: CallbackQuery):
    tg_user = callback.from_user
    username_str = f"@{tg_user.username}" if tg_user.username else "нет username"
    await notify_trainer(
        f"❓ Новый пользователь хочет зарегистрироваться:\n"
        f"Имя в TG: {tg_user.full_name}\n"
        f"Username: {username_str}\n"
        f"ID: {tg_user.id}\n\n"
        f"Добавьте его в таблицу, затем попросите написать /start."
    )
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад к списку", callback_data="back_to_list")
    await callback.message.edit_caption(
        caption="✅ Тренер уведомлён!\n\nКак только вас добавят в список — нажмите «Назад» и выберите своё имя.",
        reply_markup=b.as_markup(),
    )
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

    user_tid  = callback.from_user.id
    selected  = _multi_select.get(user_tid, {})
    confirmed = {d: tp for d, tp in selected.items() if tp != "pending"}

    text = f"🗓 {week['label']}   {week['dates']}\n\n"
    if booked:
        booked_labels = []
        for t in booked:
            short = DAY_SHORT.get(t["day"], t["day"])
            booked_labels.append(f"{short} {t['date'][:5]}")
        text += f"✅ Вы записаны: {', '.join(booked_labels)}\n\n"
    if confirmed:
        text += f"☑️ Выбрано: {len(confirmed)} дн. — нажмите «Записаться»\n"
    elif selected:
        text += "⬇️ Выберите тип тренировки:\n"
    else:
        text += "Нажмите на день чтобы выбрать:\n"

    rows_kb = []
    for t in trains:
        short      = DAY_SHORT.get(t["day"], t["day"])
        date_short = t["date"][:5]
        sel_type   = selected.get(t["date"])

        if t["booked"]:
            if t["remote"] or not t["time"]:
                label = f"✅  {short} {date_short} — 🏠 удалённо"
            else:
                label = f"✅  {short} {date_short} — 🏊 {t['time']}"
            rows_kb.append([InlineKeyboardButton(text=label, callback_data=f"day_{t['date']}")])

        elif sel_type == "pending":
            rows_kb.append([InlineKeyboardButton(
                text=f"◉  {short} {date_short} — выберите тип:",
                callback_data=f"toggle_{t['date']}",
            )])
            rows_kb.append([
                InlineKeyboardButton(text="🏊 С тренером", callback_data=f"set_pool_{t['date']}"),
                InlineKeyboardButton(text="🏠 Удалённо",  callback_data=f"set_remote_{t['date']}"),
            ])

        elif sel_type == "pool":
            rows_kb.append([InlineKeyboardButton(
                text=f"☑️  {short} {date_short} — 🏊 {t['time']}",
                callback_data=f"toggle_{t['date']}",
            )])

        elif sel_type == "remote":
            rows_kb.append([InlineKeyboardButton(
                text=f"☑️  {short} {date_short} — 🏠 удалённо",
                callback_data=f"toggle_{t['date']}",
            )])

        else:
            if _is_past_day(t["date"]):
                if t["time"]:
                    label = f"—  {short} {date_short} — {t['time']}"
                else:
                    label = f"—  {short} {date_short} — удалённо"
                rows_kb.append([InlineKeyboardButton(text=label, callback_data="noop")])
            else:
                if t["time"]:
                    label = f"🏊  {short} {date_short} — {t['time']}"
                else:
                    label = f"🏠  {short} {date_short} — удалённо"
                rows_kb.append([InlineKeyboardButton(text=label, callback_data=f"toggle_{t['date']}")])

    if confirmed:
        n    = len(confirmed)
        noun = "день" if n == 1 else "дня" if n <= 4 else "дней"
        rows_kb.append([InlineKeyboardButton(
            text=f"✅  ЗАПИСАТЬСЯ НА {n} {noun.upper()}  ✅",
            callback_data="book_selected",
        )])
    rows_kb.append([InlineKeyboardButton(text="🔄 Обновить", callback_data="refresh_schedule")])
    rows_kb.append([InlineKeyboardButton(text="◀️ Назад",    callback_data="main_menu")])

    try:
        await callback.message.edit_caption(caption=text, reply_markup=InlineKeyboardMarkup(inline_keyboard=rows_kb))
    except TelegramBadRequest:
        pass  # содержимое не изменилось — игнорируем
    try:
        await callback.answer(_toast)
    except TelegramBadRequest:
        pass  # callback устарел (>30 сек) — игнорируем

# ── Noop (прошедшие незабронированные дни) ──────────────────────
@dp.callback_query(F.data == "noop")
async def cb_noop(callback: CallbackQuery):
    await callback.answer()

# ── Обновить расписание ──────────────────────────────────────────
@dp.callback_query(F.data == "refresh_schedule")
async def cb_refresh_schedule(callback: CallbackQuery):
    _multi_select.pop(callback.from_user.id, None)
    _invalidate_bd()
    await cb_schedule(callback, "✅ Расписание обновлено")

# ── Переключение выбора дня (мультивыбор) ───────────────────────
@dp.callback_query(F.data.startswith("toggle_"))
async def cb_toggle_day(callback: CallbackQuery):
    date_str  = callback.data[7:]
    if _is_past_day(date_str):
        await callback.answer("⛔ Этот день уже прошёл")
        return
    user_tid  = callback.from_user.id
    user_name = get_user_name_by_telegram_id(user_tid)
    sel = _multi_select.setdefault(user_tid, {})

    if date_str in sel:
        del sel[date_str]
    else:
        _ensure_bd()
        trains = get_schedule_for_user(user_name or "")
        t = next((x for x in trains if x["date"] == date_str), None)
        if t and not t["time"]:
            sel[date_str] = "remote"   # remote-only день — сразу как удалённо
        else:
            sel[date_str] = "pending"  # бассейновый день — показываем выбор типа

    await cb_schedule(callback)

# ── Выбор типа тренировки для выбранного дня ────────────────────
@dp.callback_query(F.data.startswith("set_pool_"))
async def cb_set_pool(callback: CallbackQuery):
    date_str = callback.data[9:]
    _multi_select.setdefault(callback.from_user.id, {})[date_str] = "pool"
    await cb_schedule(callback)

@dp.callback_query(F.data.startswith("set_remote_"))
async def cb_set_remote(callback: CallbackQuery):
    date_str = callback.data[11:]
    _multi_select.setdefault(callback.from_user.id, {})[date_str] = "remote"
    await cb_schedule(callback)

# ── Бронирование всех выбранных дней ────────────────────────────
@dp.callback_query(F.data == "book_selected")
async def cb_book_selected(callback: CallbackQuery):
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    user_tid  = callback.from_user.id
    selected  = _multi_select.pop(user_tid, {})
    confirmed = {d: tp for d, tp in selected.items() if tp != "pending"}
    if not confirmed:
        await callback.answer("Нет выбранных дней")
        return

    _ensure_bd()
    trains = get_schedule_for_user(user_name)
    booked_dates = []

    for date_str in sorted(confirmed, key=lambda d: _parse_sheet_date(d, _now().year) or _now().date()):
        book_type = confirmed[date_str]
        t = next((x for x in trains if x["date"] == date_str), None)
        if not t or t["booked"] or _is_past_day(date_str):
            continue
        if book_type == "remote":
            ok = set_remote_booking_comment(user_name, date_str)
            if ok:
                await notify_trainer(f"🏠 {user_name} записался удалённо — {date_str}")
        else:
            ok = set_booking(user_name, date_str, True)
            if ok:
                await notify_trainer(f"🏊 {user_name} забронировал тренировку — {date_str}")
        if ok:
            booked_dates.append((date_str, book_type))
            _last_booking[user_name] = date_str
            _inactivity_notified.discard(user_name)
            save_last_booking(user_name, date_str)

    _invalidate_bd()

    if booked_dates:
        lines = "\n".join(
            f"• {'🏠' if tp == 'remote' else '🏊'} {d}"
            for d, tp in booked_dates
        )
        await callback.message.edit_caption(
            caption=f"Вы записаны на тренировки:\n{lines}",
            reply_markup=kb_main_menu(),
        )
        await check_subscription(user_name)
    else:
        await callback.message.edit_caption(
            caption="❌ Не удалось записаться. Обратитесь к тренеру.",
            reply_markup=kb_main_menu(),
        )
    await callback.answer()

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
        if not _is_finished(date_str, t["time"]):
            b.button(text="❌ Отменить запись", callback_data=f"unbook_confirm_{date_str}")
    elif not _is_past_day(date_str):
        if is_remote_day:
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

    if _is_past_day(date_str):
        await callback.message.edit_caption(caption="⛔ Этот день уже прошёл.", reply_markup=kb_main_menu())
        await callback.answer()
        return

    ok = set_remote_booking_comment(user_name, date_str)
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

    if _is_past_day(date_str):
        await callback.message.edit_caption(caption="⛔ Этот день уже прошёл.", reply_markup=kb_main_menu())
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

    text = f"📝 *{week['label']}*   _{week['dates']}_\n\nНажмите на тренировку для подробностей:"

    booked.sort(key=lambda t: _parse_sheet_date(t["date"], _now().year) or _now().date())

    rows_kb = []
    for t in booked:
        short        = DAY_SHORT.get(t["day"], t["day"])
        time_display = t["time"] if t["time"] else "удалённо"
        fmt = "✅" if _is_finished(t["date"], t["time"]) else ("🏠" if (t["remote"] or not t["time"]) else "🏊")
        rows_kb.append([InlineKeyboardButton(
            text=f"{fmt}  {short} {t['date'][:5]} — {time_display}",
            callback_data=f"booking_detail_{t['date']}",
        )])

    rows_kb.append([InlineKeyboardButton(text="◀️ Назад", callback_data="main_menu")])
    b = InlineKeyboardMarkup(inline_keyboard=rows_kb)

    await callback.message.edit_caption(caption=text, parse_mode="Markdown", reply_markup=b)
    await callback.answer()

# ── Детальная карточка брони (из Мои записи) ─────────────────────
@dp.callback_query(F.data.startswith("booking_detail_"))
async def cb_booking_detail(callback: CallbackQuery):
    date_str  = callback.data[15:]
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    _ensure_bd()
    trains = get_schedule_for_user(user_name)
    t = next((x for x in trains if x["date"] == date_str), None)
    if not t:
        await callback.message.edit_caption(caption="❌ Тренировка не найдена.")
        await callback.answer()
        return

    is_remote = t["remote"] or not t["time"]
    if is_remote:
        header = f"🏠 {t['day']} {date_str} — удалённо"
    else:
        header = f"🏊 {t['day']} {date_str} — {t['time']}"

    text = f"*{header}*\n\n"
    if t["plan"]:
        text += f"📋 *План:*\n{t['plan']}\n\n"
    else:
        text += "📋 *План:* тренер ещё не написал\n\n"
    if t["volume"]:
        text += f"📏 *Объём:* {t['volume']}\n"
    if t["comments"]:
        text += f"💬 *Комментарий:* {t['comments']}\n"

    has_plan_and_volume = bool(t["plan"] and t["volume"])

    b = InlineKeyboardBuilder()
    if _is_finished(date_str, t["time"]):
        pass  # тренировка завершена — только кнопка назад
    else:
        if has_plan_and_volume and not is_remote:
            b.button(text="🏠 Перевести на удалёнку", callback_data=f"to_remote_{date_str}")
        elif not has_plan_and_volume:
            b.button(text="❌ Отменить запись", callback_data=f"unbook_confirm_{date_str}")
        b.button(text="👥 Кто будет на тренировке", callback_data=f"participants_{date_str}")
    b.button(text="◀️ К записям", callback_data="my_booking")
    b.adjust(1)

    await callback.message.edit_caption(caption=text, parse_mode="Markdown", reply_markup=b.as_markup())
    await callback.answer()

# ── Участники тренировки ─────────────────────────────────────────
@dp.callback_query(F.data.startswith("participants_"))
async def cb_participants(callback: CallbackQuery):
    date_str  = callback.data[13:]
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.answer("❌ Вы не зарегистрированы.")
        return

    _ensure_bd()
    if not _bd_rows or not _bd_rows[0]:
        await callback.answer("Расписание недоступно.")
        return

    headers = _bd_rows[1]
    col = {h: i for i, h in enumerate(headers)}
    pool_list   = []
    remote_list = []

    for row in _bd_rows[2:]:
        if not row or len(row) <= col.get("Date", 1):
            continue
        if row[col["Date"]].strip() != date_str:
            continue
        if row[col.get("Booked", 7)].upper() != "TRUE":
            continue
        name = row[col["User"]].strip()
        if row[col.get("Remote", 6)].lower() == "yes":
            remote_list.append(name)
        else:
            pool_list.append(name)

    text = f"👥 *Тренировка {date_str}*\n\n"
    if pool_list:
        text += "🏊 *С тренером:*\n" + "\n".join(f"• {n}" for n in pool_list) + "\n\n"
    if remote_list:
        text += "🏠 *Удалённо:*\n" + "\n".join(f"• {n}" for n in remote_list) + "\n"
    if not pool_list and not remote_list:
        text += "Пока никто не записался."

    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data=f"booking_detail_{date_str}")

    try:
        await callback.message.edit_caption(caption=text, parse_mode="Markdown", reply_markup=b.as_markup())
    except TelegramBadRequest:
        pass
    await callback.answer()

# ── Перевод на удалёнку ──────────────────────────────────────────
@dp.callback_query(F.data.startswith("to_remote_"))
async def cb_to_remote(callback: CallbackQuery):
    date_str  = callback.data[10:]
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return

    ok = set_remote_booking_comment(user_name, date_str)
    if ok:
        _invalidate_bd()
        await callback.message.edit_caption(
            caption=f"🏠 Тренировка {date_str} переведена на удалёнку.",
            reply_markup=kb_main_menu(),
        )
        await notify_trainer(f"🏠 {user_name} перешёл на удалённую тренировку — {date_str}")
    else:
        await callback.message.edit_caption(caption="❌ Не удалось изменить. Обратитесь к тренеру.")
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

    if _is_finished(date_str, t["time"] if t else ""):
        b = InlineKeyboardBuilder()
        b.button(text="◀️ Назад", callback_data="my_booking")
        await callback.message.edit_caption(
            caption="❌ Отменить нельзя — тренировка уже прошла.",
            reply_markup=b.as_markup(),
        )
        await callback.answer()
        return
    time_display = t["time"] if t and t["time"] else "удалённо"
    day_name = t["day"] if t else date_str
    has_plan_and_volume = bool(t and t["plan"] and t["volume"])
    is_remote = bool(t and (t["remote"] or not t["time"]))

    b = InlineKeyboardBuilder()
    if has_plan_and_volume and not is_remote:
        b.button(text="🏠 Перевести на удалёнку", callback_data=f"to_remote_{date_str}")
        b.button(text="◀️ Оставить",              callback_data="my_booking")
        caption = (
            f"*{day_name} {date_str} — {time_display}*\n\n"
            f"❌ Отменить нельзя — тренер уже написал план и объём.\n"
            f"Можно перевести тренировку на удалёнку."
        )
    elif has_plan_and_volume and is_remote:
        b.button(text="◀️ Назад", callback_data="my_booking")
        caption = (
            f"*{day_name} {date_str} — удалённо*\n\n"
            f"❌ Отменить нельзя — тренер уже написал план и объём."
        )
    else:
        b.button(text="✅ Да, отменить",   callback_data=f"unbook_{date_str}")
        b.button(text="◀️ Нет, оставить", callback_data="my_booking")
        caption = (
            f"Вы уверены что хотите отменить запись?\n\n"
            f"*{day_name} {date_str} — {time_display}*"
        )
    b.adjust(1)

    await callback.message.edit_caption(caption=caption, parse_mode="Markdown", reply_markup=b.as_markup())
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

    ok = cancel_booking(user_name, date_str)

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
    grade            = get_user_grade(user_name)

    # Метры за текущую неделю — сумма из BD по записанным тренировкам
    try:
        _ensure_bd()
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

    dnf_grade  = get_user_dnf_grade(user_name)
    grade_parts = []
    if grade:     grade_parts.append(f"🏊 {grade}")
    if dnf_grade: grade_parts.append(f"🤿 {dnf_grade}")
    grade_line = "🏅 *Уровни:* " + "  |  ".join(grade_parts) + "\n\n" if grade_parts else ""
    await callback.message.edit_caption(caption=
        f"👤 *Профиль — {user_name}*\n\n"
        f"{grade_line}"
        f"🏊 Метров на этой неделе:\n*{meters_week}*\n\n"
        f"📅 Метров в {2026} году:\n*{meters_year}*\n\n"
        f"📅 Метров в {2025} году:\n*{meters_last_year}*\n"
        f"{compare}",
        parse_mode="Markdown",
        reply_markup=b.as_markup(),
    )
    try:
        await callback.answer()
    except TelegramBadRequest:
        pass

# ── Оценка состояния ─────────────────────────────────────────────
@dp.callback_query(F.data.startswith("state_"))
async def cb_state(callback: CallbackQuery):
    parts      = callback.data.split("_")
    date_str   = parts[1]
    idx        = int(parts[2])
    state_value = STATES[idx]
    user_name  = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.answer("❌ Вы не зарегистрированы.")
        return
    ok = set_state(user_name, date_str, state_value)
    if ok:
        await callback.message.edit_text(
            f"✅ Состояние сохранено: *{state_value}*\n\nСпасибо, {user_name}!",
            parse_mode="Markdown",
        )
        await notify_trainer(f"📊 {user_name} оценил тренировку {date_str}: {state_value}")
        await callback.answer()
    else:
        await callback.answer("❌ Не удалось сохранить. Попробуй позже.")

# ── Просмотр плана из напоминания ───────────────────────────────
@dp.callback_query(F.data.startswith("view_plan_"))
async def cb_view_plan(callback: CallbackQuery):
    date_str  = callback.data[10:]
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.answer("❌ Вы не зарегистрированы.")
        return

    _ensure_bd()
    trains = get_schedule_for_user(user_name)
    t = next((x for x in trains if x["date"] == date_str), None)
    if not t:
        await callback.answer("❌ Тренировка не найдена.")
        return

    text = f"🏊 *{t['day']} {date_str} — {t['time']}*\n\n"
    if t["plan"]:
        text += f"📋 *План:*\n{t['plan']}\n\n"
    else:
        text += "📋 *План:* тренер ещё не написал — ожидайте\n\n"
    if t["volume"]:
        text += f"📏 *Объём:* {t['volume']}\n"
    if t["comments"]:
        text += f"💬 *Комментарий:* {t['comments']}\n"

    try:
        await callback.message.edit_text(text, parse_mode="Markdown")
    except TelegramBadRequest:
        pass
    await callback.answer()

# ── Мой уровень — хаб ────────────────────────────────────────────
@dp.callback_query(F.data == "my_level")
async def cb_my_level(callback: CallbackQuery):
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return
    b = InlineKeyboardBuilder()
    b.button(text="🏊 Плавание", callback_data="swim_levels")
    b.button(text="🤿 DNF/DYN",  callback_data="dnf_levels")
    b.button(text="◀️ Назад",    callback_data="main_menu")
    b.adjust(2, 1)
    await callback.message.edit_caption(
        caption=f"🎯 *Мой уровень — {user_name}*\n\nВыберите дисциплину:",
        parse_mode="Markdown",
        reply_markup=b.as_markup(),
    )
    await callback.answer()

# ── Плавание — список уровней ─────────────────────────────────────
@dp.callback_query(F.data == "swim_levels")
async def cb_swim_levels(callback: CallbackQuery):
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return
    current     = _known_grades.get(user_name, "")
    current_idx = SWIM_GRADE_ORDER.index(current) if current in SWIM_GRADE_ORDER else -1
    b = InlineKeyboardBuilder()
    for i, grade in enumerate(SWIM_GRADE_ORDER):
        if i < current_idx:
            label = f"✅  {grade}"
        elif i == current_idx:
            label = f"📍  {grade}"
        elif i == current_idx + 1:
            label = f"→  {grade}"
        else:
            label = f"     {grade}"
        b.button(text=label, callback_data=f"swim_lvl_{i}")
    b.button(text="◀️ Назад", callback_data="my_level")
    b.adjust(1)
    text = f"🏊 *Плавание — {user_name}*\n\n"
    if current:
        text += f"Текущий уровень: *{current}*\n"
        if current_idx + 1 < len(SWIM_GRADE_ORDER):
            text += f"Следующий: *{SWIM_GRADE_ORDER[current_idx + 1]}*\n"
    else:
        text += "Уровень ещё не присвоен.\n"
    text += "\nНажмите на уровень чтобы посмотреть условия:"
    await callback.message.edit_caption(caption=text, parse_mode="Markdown", reply_markup=b.as_markup())
    await callback.answer()

# ── DNF/DYN — список уровней ──────────────────────────────────────
@dp.callback_query(F.data == "dnf_levels")
async def cb_dnf_levels(callback: CallbackQuery):
    user_name = get_user_name_by_telegram_id(callback.from_user.id)
    if not user_name:
        await callback.message.edit_caption(caption="❌ Вы не зарегистрированы.")
        await callback.answer()
        return
    current     = _known_dnf_grades.get(user_name, "")
    current_idx = DNF_GRADE_ORDER.index(current) if current in DNF_GRADE_ORDER else -1
    b = InlineKeyboardBuilder()
    for i, grade in enumerate(DNF_GRADE_ORDER):
        if i < current_idx:
            label = f"✅  {grade}"
        elif i == current_idx:
            label = f"📍  {grade}"
        elif i == current_idx + 1:
            label = f"→  {grade}"
        else:
            label = f"     {grade}"
        b.button(text=label, callback_data=f"dnf_lvl_{i}")
    b.button(text="◀️ Назад", callback_data="my_level")
    b.adjust(1)
    text = f"🤿 *DNF/DYN — {user_name}*\n\n"
    if current:
        text += f"Текущий уровень: *{current}*\n"
        if current_idx + 1 < len(DNF_GRADE_ORDER):
            text += f"Следующий: *{DNF_GRADE_ORDER[current_idx + 1]}*\n"
    else:
        text += "Уровень ещё не присвоен.\n"
    text += "\nНажмите на уровень чтобы посмотреть условия:"
    await callback.message.edit_caption(caption=text, parse_mode="Markdown", reply_markup=b.as_markup())
    await callback.answer()

# ── Детали уровня плавания ────────────────────────────────────────
@dp.callback_query(F.data.startswith("swim_lvl_"))
async def cb_swim_level_detail(callback: CallbackQuery):
    idx = int(callback.data[9:])
    if not (0 <= idx < len(SWIM_GRADE_ORDER)):
        await callback.answer()
        return
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data="swim_levels")
    await callback.message.edit_caption(
        caption=SWIM_GRADE_INFO[SWIM_GRADE_ORDER[idx]],
        parse_mode="Markdown",
        reply_markup=b.as_markup(),
    )
    await callback.answer()

# ── Детали уровня DNF/DYN ─────────────────────────────────────────
@dp.callback_query(F.data.startswith("dnf_lvl_"))
async def cb_dnf_level_detail(callback: CallbackQuery):
    idx = int(callback.data[8:])
    if not (0 <= idx < len(DNF_GRADE_ORDER)):
        await callback.answer()
        return
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data="dnf_levels")
    await callback.message.edit_caption(
        caption=DNF_GRADE_INFO[DNF_GRADE_ORDER[idx]],
        parse_mode="Markdown",
        reply_markup=b.as_markup(),
    )
    await callback.answer()

# ── Главное меню ─────────────────────────────────────────────────
@dp.callback_query(F.data == "main_menu")
async def cb_main_menu(callback: CallbackQuery):
    _multi_select.pop(callback.from_user.id, None)
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
            now = _now()
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
            now = _now()
            if now.hour != 10 or now.minute != 0:
                continue
            for user_name, last_date_str in list(_last_booking.items()):
                if user_name in _inactivity_notified:
                    continue
                try:
                    last_date = datetime.strptime(last_date_str, "%d.%m.%Y").date()
                except ValueError:
                    try:
                        last_date = datetime.strptime(last_date_str + f".{now.year}", "%d.%m.%Y").date()
                    except Exception:
                        continue
                days_gone = (now.date() - last_date).days
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
            now = _now()
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
                    if not t["booked"] or not t["time"] or not _is_today(t["date"], now):
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
                            tid = _tid_cache.get(user_name)
                            if not tid:
                                continue
                            rb = InlineKeyboardBuilder()
                            rb.button(text="📋 Посмотреть план тренировки", callback_data=f"view_plan_{t['date']}")
                            await bot.send_message(
                                tid,
                                f"⏰ {user_name}, напоминание!\n\n"
                                f"Через ~2 часа тренировка в {t['time']} 🏊\n"
                                f"{t['day']}, {t['date']}",
                                reply_markup=rb.as_markup(),
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
                        tid = _tid_cache.get(user_name)
                        if tid:
                            pb = InlineKeyboardBuilder()
                            pb.button(text="📋 Посмотреть план", callback_data=f"view_plan_{t['date']}")
                            await bot.send_message(
                                tid,
                                f"📋 {user_name}, тренер написал план тренировки, можно ознакомиться!\n\n"
                                f"{t['day']} {t['date']}",
                                reply_markup=pb.as_markup(),
                            )
        except Exception as e:
            print(f"plan_checker ошибка: {e}")


async def state_checker():
    """Через 2 часа после тренировки спрашивает ученика о состоянии."""
    global _state_notified, _state_date
    while True:
        await asyncio.sleep(60)
        try:
            now = _now()
            today_str = now.strftime("%d.%m.%Y")

            if today_str != _state_date:
                _state_notified = set()
                _state_date = today_str

            _ensure_bd()
            if not _bd_rows or not _bd_rows[0]:
                continue

            for user_name in USER_COLUMNS:
                trains = get_schedule_for_user(user_name)
                for t in trains:
                    if not t["booked"] or not t["time"] or not _is_today(t["date"], now):
                        continue
                    key = f"{user_name}|{t['date']}"
                    if key in _state_notified:
                        continue
                    try:
                        h, m = map(int, t["time"].split(":"))
                        train_dt  = now.replace(hour=h, minute=m, second=0, microsecond=0)
                        delta_min = (now - train_dt).total_seconds() / 60
                        if 115 <= delta_min <= 140:
                            _state_notified.add(key)
                            tid = _tid_cache.get(user_name)
                            if not tid:
                                continue
                            b = InlineKeyboardBuilder()
                            for idx, state in enumerate(STATES[1:], 1):
                                b.button(text=state, callback_data=f"state_{t['date']}_{idx}")
                            b.adjust(1)
                            await bot.send_message(
                                tid,
                                f"🏊 {user_name}, как прошла тренировка {t['date']} в {t['time']}?\n\n"
                                f"Оцени своё состояние:",
                                reply_markup=b.as_markup(),
                            )
                    except Exception:
                        continue
        except Exception as e:
            print(f"state_checker ошибка: {e}")

async def grade_checker():
    """Каждые 5 минут проверяет изменение грейдов в строке 12."""
    while True:
        await asyncio.sleep(300)
        try:
            row = get_source_sheet().row_values(12)
            for name, col_letter in USER_COLUMNS.items():
                col_idx = ord(col_letter) - ord("A")
                raw = str(row[col_idx]).strip() if col_idx < len(row) else ""
                new_swim = _parse_grade(raw)
                new_dnf  = _parse_dnf_grade(raw)
                old_swim = _known_grades.get(name)
                old_dnf  = _known_dnf_grades.get(name)

                if old_swim is None:
                    _known_grades[name] = new_swim
                elif new_swim != old_swim:
                    _known_grades[name] = new_swim
                    save_grade_history(name, new_swim)
                    if new_swim:
                        await notify_user(
                            name,
                            f"🏅 Поздравляем, {name}!\n\nУ вас новый уровень плавания: *{new_swim}*",
                            parse_mode="Markdown",
                        )

                if old_dnf is None:
                    _known_dnf_grades[name] = new_dnf
                elif new_dnf != old_dnf:
                    _known_dnf_grades[name] = new_dnf
                    if new_dnf:
                        await notify_user(
                            name,
                            f"🤿 Поздравляем, {name}!\n\nУ вас новый уровень DNF/DYN: *{new_dnf}*",
                            parse_mode="Markdown",
                        )
        except Exception as e:
            print(f"grade_checker ошибка: {e}")

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
    print("Бот запущен 🚀 v2.0")
    load_all_tids()
    load_last_bookings()
    load_grades()
    get_grades_sheet()  # создаёт лист Grades если не существует
    asyncio.create_task(week_watcher())
    asyncio.create_task(grade_checker())
    asyncio.create_task(weekly_report())
    asyncio.create_task(inactivity_checker())
    asyncio.create_task(training_reminder())
    asyncio.create_task(plan_checker())
    asyncio.create_task(state_checker())
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
