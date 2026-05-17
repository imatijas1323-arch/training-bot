"""
bot.py — Swimming Training Bot
"""

import asyncio
import html
import os
import re
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, CallbackQuery, ReplyKeyboardMarkup, KeyboardButton, FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.filters import CommandStart, Command
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.exceptions import TelegramBadRequest

load_dotenv()

# ═══════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ═══════════════════════════════════════════════════════════════

TOKEN          = os.getenv("BOT_TOKEN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
TRAINER_ID       = int(os.getenv("TRAINER_ID", "0"))
TRAINER_PASSWORD = os.getenv("TRAINER_PASSWORD", "")
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
    "JUNIOR ⭑",
    "NOT BAD ⭑⭑⭑",
    "NOT BAD ★★",
    "NOT BAD ☆",
    "LEADER ⭑⭑⭑",
    "LEADER ★★",
    "LEADER ☆",
    "ELITE ✪",
    "PRO ✪",
]

SWIM_GRADE_INFO = {
    "JUNIOR ⭑": (
        "🏊 *JUNIOR ⭑*\n\n"
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
    "ELITE ✪": (
        "🏊 *ELITE ✪*\n\n"
        "📏 4000м вольным стилем\n"
        "⏱ Темп 50м: 29 с (М) / 33 с (Ж)\n"
        "⏱ 1000м до 16 мин (М) / 19 мин (Ж)"
    ),
    "PRO ✪": (
        "🏊 *PRO ✪*\n\n"
        "Высший уровень — соревновательные ранги и титулы."
    ),
}

DNF_GRADE_ORDER = ["DNF/DYN ★", "DNF/DYN ★★", "DNF/DYN ★★★"]

DNF_GRADE_INFO = {
    "DNF/DYN ★": (
        "🤿 *DNF/DYN ★*\n\n"
        "🫁 DNF (без ласт): 25 м\n"
        "🦈 DYN (с ластами): 35 м"
    ),
    "DNF/DYN ★★": (
        "🤿 *DNF/DYN ★★*\n\n"
        "🫁 DNF (без ласт): 35 м\n"
        "🦈 DYN (с ластами): 50 м"
    ),
    "DNF/DYN ★★★": (
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

def get_sub_sheet():
    return ss.worksheet("абонемент")

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

def get_grade_current_sheet():
    for ws in ss.worksheets():
        if ws.title == "GradeCurrent":
            return ws
    ws = ss.add_worksheet(title="GradeCurrent", rows=50, cols=3)
    ws.update([["Имя", "Плавание", "DNF/DYN"]], "A1")
    return ws

# ═══════════════════════════════════════════════════════════════
# КЭШ
# ═══════════════════════════════════════════════════════════════

_user_cache:          dict[int, str] = {}   # telegram_id → имя
_tid_cache:           dict[str, int] = {}   # имя → telegram_id
_bd_rows:             list           = []   # строки из листа BD
_bd_ts:               float          = 0.0  # время последней синхронизации
_week_days_cache:     list           = []   # кэш tr_get_current_week_days()
_week_days_ts:        float          = 0.0  # TTL кэша дней недели
_week_marker_row:       int  = -1    # строка с "текущая неделя"
_prev_week_row:         int  = -1    # строка последней закрытой недели
_pending_open_row:      int  = -1    # развёрнутая но ещё не открытая неделя
_viewing_prev_week:     bool = False # тренер просматривает прошлую неделю
_week_collapsed:        bool = False # текущая неделя свёрнута в панели тренера
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
_authenticated_trainers: set[int]              = set()  # trainer IDs прошедших авторизацию
_trainer_state:          dict[int, dict]       = {}     # {user_id: {action, ...}} для многошаговых потоков
_trainer_view_as:        dict[int, str]        = {}     # trainer_id → имя ученика (режим просмотра)
_tr_expanded:            dict[int, str]        = {}     # trainer_id → развёрнутая дата в разделе Тренировки
_tr_plan_sel:            dict[int, dict]       = {}     # trainer_id → {date_str: [names]} мультивыбор для плана
_tr_viewing_prev:        set                   = set()  # trainer_id → просматривает прошлую неделю
_sub_pending:            dict[int, dict]       = {}     # trainer_id → {name: pending_delta}
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
    global _bd_ts, _week_days_ts
    _bd_ts = 0.0
    _week_days_ts = 0.0

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
    global _bd_ts
    if (time.time() - _bd_ts) > BD_TTL:
        sync_bd()  # sync_bd сам обновляет _bd_rows и _bd_ts

# ═══════════════════════════════════════════════════════════════
# ПОЛЬЗОВАТЕЛИ
# ═══════════════════════════════════════════════════════════════

def get_user_name_by_telegram_id(telegram_id: int) -> str | None:
    if telegram_id in _trainer_view_as:
        return _trainer_view_as[telegram_id]
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
    """Нормализует любой вариант написания к каноничному грейду из SWIM_GRADE_ORDER."""
    for keyword, symbol in [
        ("NOT BAD", "⭑⭑⭑"), ("NOT BAD", "★★"), ("NOT BAD", "☆"),
        ("LEADER",  "⭑⭑⭑"), ("LEADER",  "★★"), ("LEADER",  "☆"),
    ]:
        if keyword in text and symbol in text:
            return f"{keyword} {symbol}"
    # одна ★ без второй → высший подуровень (☆ / ⭑)
    has_single_star = "★" in text and "★★" not in text
    if "NOT BAD" in text: return "NOT BAD ☆" if has_single_star else "NOT BAD ⭑⭑⭑"
    if "LEADER"  in text: return "LEADER ☆"  if has_single_star else "LEADER ⭑⭑⭑"
    if "PRO"     in text: return "PRO ✪"
    if "ELITE"   in text: return "ELITE ✪"
    if "JUNIOR"  in text: return "JUNIOR ⭑"
    return ""

def _extract_dnf_grade(text: str) -> str:
    """Извлекает DNF/DYN грейд из текста (правая часть после '/')."""
    text = text.strip()
    if "★★★" in text: return "DNF/DYN ★★★"
    if "★★"  in text: return "DNF/DYN ★★"
    if "★"   in text: return "DNF/DYN ★"
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
        ws = get_grade_current_sheet()
        rows = ws.get_all_values()
        for row in rows[1:]:
            if row and row[0].strip() == user_name:
                return row[1].strip() if len(row) > 1 else ""
    except Exception:
        pass
    return ""

def get_user_dnf_grade(user_name: str) -> str:
    if user_name in _known_dnf_grades:
        return _known_dnf_grades[user_name]
    try:
        ws = get_grade_current_sheet()
        rows = ws.get_all_values()
        for row in rows[1:]:
            if row and row[0].strip() == user_name:
                return row[2].strip() if len(row) > 2 else ""
    except Exception:
        pass
    return ""

def load_grades():
    """Загружает грейды из листа GradeCurrent в кэш."""
    try:
        ws = get_grade_current_sheet()
        rows = ws.get_all_values()
        grades_sheet = get_grades_sheet()
        existing = grades_sheet.get_all_values()
        users_with_entry = {r[0] for r in existing[1:] if r}
        date_str = _now().strftime("%d.%m.%Y")
        rows_to_add = []
        for row in rows[1:]:
            if not row or not row[0].strip():
                continue
            name = row[0].strip()
            swim = _extract_swim_grade(row[1].strip() if len(row) > 1 else "")
            dnf  = _extract_dnf_grade(row[2].strip() if len(row) > 2 else "")
            _known_grades[name] = swim
            _known_dnf_grades[name] = dnf
            if swim and name not in users_with_entry:
                rows_to_add.append([name, swim, date_str])
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
    global _bd_rows, _bd_ts
    try:
        src  = get_source_sheet()
        data = src.get_all_values()
    except Exception as e:
        print(f"sync_bd() ошибка чтения: {e}")
        return

    # Находим строку с маркером текущей недели
    week_bron_idx = None
    for i, row in enumerate(data):
        last = str(row[-1]).strip().lower() if row else ""
        if "текущ" in last:
            week_bron_idx = i
            break
    if week_bron_idx is None:
        get_bd_sheet().clear()
        _bd_rows = []
        _bd_ts   = time.time()
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

    # Пишем в BD одним запросом + обновляем in-memory кэш без повторного чтения
    header_row = ["User", "Date", "Day", "Time", "Plan", "Volume", "Remote", "Booked", "Comments"]
    all_rows = [[week_line], header_row] + records
    try:
        bd = get_bd_sheet()
        bd.clear()
        bd.update(all_rows, "A1")
    except Exception as e:
        print(f"sync_bd() ошибка записи в BD: {e}")
        return
    _bd_rows = all_rows
    _bd_ts   = time.time()

# ═══════════════════════════════════════════════════════════════
# ТРЕНЕР — ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════════

def is_trainer(user_id: int) -> bool:
    return user_id == TRAINER_ID or user_id in _authenticated_trainers

_gauth_session: AuthorizedSession | None = None

def _get_gauth_session() -> AuthorizedSession:
    global _gauth_session
    if _gauth_session is None:
        _gauth_session = AuthorizedSession(creds)
    return _gauth_session

def get_sheet_row_groups() -> list:
    """Возвращает список rowGroups для листа 2026 через Sheets API v4."""
    try:
        import json as _json
        session = _get_gauth_session()
        url  = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}"
        resp = session.get(url, params={"fields": "sheets(rowGroups,properties.sheetId)"})
        data = _json.loads(resp.content)
        src_id = get_source_sheet().id
        for sheet in data.get("sheets", []):
            if sheet.get("properties", {}).get("sheetId") == src_id:
                return sheet.get("rowGroups", [])
    except Exception as e:
        print(f"get_sheet_row_groups error: {e}")
    return []

def toggle_week_collapse(bron_row_0based: int, collapse: bool) -> bool:
    """Сворачивает или разворачивает группу строк недели в таблице."""
    try:
        groups   = get_sheet_row_groups()
        sheet_id = get_source_sheet().id
        target   = None
        best_size = float("inf")
        for g in groups:
            r = g.get("range", {})
            if r.get("sheetId") != sheet_id:
                continue
            start = r.get("startIndex", 0)
            end   = r.get("endIndex",   0)
            size  = end - start
            if start <= bron_row_0based < end and size < best_size:
                target    = g
                best_size = size
        if not target:
            print(f"toggle_week_collapse: группа не найдена для строки {bron_row_0based}")
            return False
        ss.batch_update({"requests": [{
            "updateDimensionGroup": {
                "dimensionGroup": {
                    "range":     target["range"],
                    "depth":     target.get("depth", 1),
                    "collapsed": collapse,
                },
                "fields": "collapsed",
            }
        }]})
        return True
    except Exception as e:
        print(f"toggle_week_collapse error: {e}")
        return False

DAY_SHORT_RU = {
    "Monday": "Пн", "Tuesday": "Вт", "Wednesday": "Ср",
    "Thursday": "Чт", "Friday": "Пт", "Saturday": "Сб", "Sunday": "Вс",
}

def tr_get_current_week_days(override_row: int = -1) -> list[dict]:
    """Читает 7 дней недели прямо из листа 2026. override_row — если нужна не текущая неделя."""
    global _week_days_cache, _week_days_ts
    row = override_row if override_row != -1 else _week_marker_row
    if row == -1:
        return []
    # Кэш только для текущей недели (override = прошлая/следующая — не кэшируем)
    if override_row == -1 and _week_days_cache and (time.time() - _week_days_ts) < BD_TTL:
        return _week_days_cache
    data = get_source_sheet().get_all_values()
    result = []
    row_idx = row
    for _ in range(7):
        if row_idx + 3 >= len(data):
            break
        plan_row  = data[row_idx + 1]
        vol_row   = data[row_idx + 2]
        com_row   = data[row_idx + 3]
        day_ru    = str(plan_row[1]).strip().lower() if len(plan_row) > 1 else ""
        day_en    = DAY_MAP_SYNC.get(day_ru, day_ru.capitalize())
        day_short = DAY_SHORT_RU.get(day_en, day_en[:2])
        date_str  = str(com_row[1]).strip()  if len(com_row)  > 1 else ""
        time_raw  = str(vol_row[1]).strip()  if len(vol_row)  > 1 else ""
        is_no_train = "нет тренировки" in time_raw.lower()
        is_remote   = "удал" in time_raw.lower()
        time_clean  = ""
        if not is_no_train and not is_remote and time_raw:
            time_clean = re.sub(r"(\d{1,2})\s+(\d{2})", r"\1:\2", time_raw)
        result.append({
            "day_short":   day_short,
            "date":        date_str,
            "time":        time_clean,
            "is_no_train": is_no_train,
            "is_remote":   is_remote,
        })
        row_idx += 5
    if override_row == -1:
        _week_days_cache = result
        _week_days_ts    = time.time()
    return result

def load_week_rows():
    """Загружает строки текущей и прошлой недели, а также состояние collapse из таблицы."""
    global _week_marker_row, _prev_week_row, _week_collapsed
    try:
        data = get_source_sheet().get_all_values()
        for i, row in enumerate(data):
            last = str(row[-1]).strip().lower() if row else ""
            if "текущ" in last and _week_marker_row == -1:
                _week_marker_row = i
            elif "прошед" in last or "прошл" in last:
                _prev_week_row = i  # берём последнюю прошедшую (максимальная строка)
        # Проверяем реальное состояние collapse текущей недели в таблице
        if _week_marker_row != -1:
            sheet_id = get_source_sheet().id
            groups   = get_sheet_row_groups()
            best_size = float("inf")
            target = None
            for g in groups:
                r = g.get("range", {})
                if r.get("sheetId") != sheet_id:
                    continue
                start = r.get("startIndex", 0)
                end   = r.get("endIndex",   0)
                size  = end - start
                if start <= _week_marker_row < end and size < best_size:
                    target    = g
                    best_size = size
            if target:
                _week_collapsed = bool(target.get("collapsed", False))
        print(f"Строки недель: текущая={_week_marker_row}, прошлая={_prev_week_row}, свёрнута={_week_collapsed}")
    except Exception as e:
        print(f"load_week_rows ошибка: {e}")

def tr_find_next_week_row() -> int:
    """Находит строку (0-based) первого дня следующей недели через группы строк."""
    base = _week_marker_row if _week_marker_row != -1 else _prev_week_row
    if base == -1:
        return -1
    try:
        sheet_id = get_source_sheet().id
        groups   = get_sheet_row_groups()
        # Берём группы глубины 3 (недели), у которых startIndex > base, сортируем
        week_groups = sorted(
            [g for g in groups
             if g.get("range", {}).get("sheetId") == sheet_id
             and g.get("depth", 0) == 3
             and g.get("range", {}).get("startIndex", 0) > base],
            key=lambda g: g["range"]["startIndex"]
        )
        if week_groups:
            return week_groups[0]["range"]["startIndex"]
    except Exception as e:
        print(f"tr_find_next_week_row группы: {e}")
    # Fallback: поиск по строкам
    data = get_source_sheet().get_all_values()
    for i in range(base + 30, min(base + 55, len(data))):
        if len(data[i]) <= 3:
            continue
        d_val = str(data[i][3]).strip().upper()
        last  = str(data[i][-1]).strip().lower()
        if d_val in ("TRUE", "FALSE") and "текущ" not in last and "прошед" not in last and "прошл" not in last:
            return i
    return -1

def tr_close_current_week() -> bool:
    global _prev_week_row, _week_marker_row, _week_collapsed
    if _week_marker_row == -1:
        return False
    _prev_week_row = _week_marker_row
    src  = get_source_sheet()
    data = src.get_all_values()
    row  = data[_week_marker_row]
    last_col = len(row)
    src.update(gspread.utils.rowcol_to_a1(_week_marker_row + 1, last_col), [["прошедшая неделя"]])
    toggle_week_collapse(_week_marker_row, True)
    _week_marker_row = -1
    _week_collapsed  = False
    _invalidate_bd()
    return True

def tr_expand_next_week() -> int:
    """Разворачивает следующую неделю в таблице без установки маркера. Возвращает строку или -1."""
    global _pending_open_row
    row = tr_find_next_week_row()
    if row == -1:
        return -1
    toggle_week_collapse(row, False)
    _pending_open_row = row
    return row

def tr_confirm_open_week() -> bool:
    """Ставит маркер 'текущая неделя' для _pending_open_row. Уведомления — отдельно."""
    global _week_marker_row, _pending_open_row, _week_collapsed, _week_session_notified
    if _pending_open_row == -1:
        return False
    row  = _pending_open_row
    src  = get_source_sheet()
    data = src.get_all_values()
    nxt_last = len(data[row])
    src.update(gspread.utils.rowcol_to_a1(row + 1, nxt_last), [["текущая неделя"]])
    _week_marker_row      = row
    _pending_open_row     = -1
    _week_collapsed       = False
    _week_session_notified = True  # блокируем авто-уведомление week_watcher
    _invalidate_bd()
    return True

def tr_cancel_expand_week() -> bool:
    """Сворачивает обратно�� развёрнутую но ещё не открытую неделю."""
    global _pending_open_row
    if _pending_open_row == -1:
        return False
    toggle_week_collapse(_pending_open_row, True)
    _pending_open_row = -1
    return True

def tr_set_time(date_str: str, time_str: str) -> bool:
    rows = find_date_rows(date_str)
    if not rows:
        return False
    get_source_sheet().update(f"B{rows['vol']}", [[time_str]])
    _invalidate_bd()
    return True

def tr_cancel_day(date_str: str) -> bool:
    return tr_set_time(date_str, "нет тренировки")

def _sub_col(user_name: str) -> str:
    """Колонка в листе 'абонемент': D→B, E→C, ... (2026 смещён на 2)"""
    return chr(ord(USER_COLUMNS[user_name]) - 2)

def _fmt_sub(val) -> str:
    """91.0 → '91', 91.5 → '91.5'"""
    try:
        f = float(str(val).replace(",", "."))
        return str(int(f)) if f == int(f) else f"{f:g}"
    except Exception:
        return str(val)

def tr_get_subscription(user_name: str) -> float:
    try:
        col = _sub_col(user_name)
        val = get_sub_sheet().acell(f"{col}5").value
        return float(str(val or "0").replace(",", "."))
    except Exception as e:
        print(f"tr_get_subscription({user_name}) ошибка: {e}")
        return 0.0

def tr_add_subscription(user_name: str, amount: float) -> bool:
    """Добавляет строку оплаты в лист 'абонемент'; формулы пересчитают остаток сами."""
    try:
        ws = get_sub_sheet()
        col = _sub_col(user_name)
        col_idx = ord(col) - ord("A") + 1
        col_a = ws.col_values(1)
        next_row = max(len(col_a) + 1, 9)
        ws.update_cell(next_row, 1, _now().strftime("%d.%m.%Y"))
        ws.update_cell(next_row, col_idx, amount)
        return True
    except Exception as e:
        print(f"tr_add_subscription({user_name}) ошибка: {e}")
        return False

def tr_set_grade(user_name: str, swim: str = None, dnf: str = None) -> bool:
    try:
        ws   = get_grade_current_sheet()
        rows = ws.get_all_values()
        for i, row in enumerate(rows):
            if row and row[0].strip() == user_name:
                if swim is not None:
                    ws.update(f"B{i+1}", [[swim]])
                    _known_grades[user_name] = swim
                if dnf is not None:
                    ws.update(f"C{i+1}", [[dnf]])
                    _known_dnf_grades[user_name] = dnf
                return True
        swim_v = swim if swim is not None else _known_grades.get(user_name, "")
        dnf_v  = dnf  if dnf  is not None else _known_dnf_grades.get(user_name, "")
        ws.append_row([user_name, swim_v, dnf_v])
        if swim is not None: _known_grades[user_name] = swim
        if dnf  is not None: _known_dnf_grades[user_name] = dnf
        return True
    except Exception as e:
        print(f"tr_set_grade ошибка: {e}")
        return False

def tr_get_week_stats() -> dict:
    _ensure_bd()
    if len(_bd_rows) < 2:
        return {}
    hdr = _bd_rows[1]
    col = {h: i for i, h in enumerate(hdr)}
    stats: dict[str, int] = {}
    for row in _bd_rows[2:]:
        if not row:
            continue
        name = str(row[col.get("User", 0)]).strip() if col.get("User", 0) < len(row) else ""
        booked = str(row[col.get("Booked", 7)]).upper() if col.get("Booked", 7) < len(row) else ""
        if name and booked == "TRUE":
            stats[name] = stats.get(name, 0) + 1
    return stats

def tr_get_student_bookings(user_name: str) -> list:
    _ensure_bd()
    if len(_bd_rows) < 2:
        return []
    hdr = _bd_rows[1]
    col = {h: i for i, h in enumerate(hdr)}
    result = []
    for row in _bd_rows[2:]:
        if not row:
            continue
        name = str(row[col.get("User", 0)]).strip() if col.get("User", 0) < len(row) else ""
        booked = str(row[col.get("Booked", 7)]).upper() if col.get("Booked", 7) < len(row) else ""
        if name == user_name and booked == "TRUE":
            result.append(row)
    return result

def tr_get_training_week_data() -> dict:
    """Данные по дням текущей недели: {date: {day_short, time, is_no_train, is_remote, participants}}"""
    _ensure_bd()
    if len(_bd_rows) < 2:
        return {}
    hdr = _bd_rows[1]
    col = {h: i for i, h in enumerate(hdr)}
    days_ordered = tr_get_current_week_days()
    result = {}
    for day_info in days_ordered:
        date_str = day_info["date"]
        if not date_str:
            continue
        participants = []
        for row in _bd_rows[2:]:
            if not row or col.get("Date", 1) >= len(row):
                continue
            if row[col["Date"]].strip() != date_str:
                continue
            if str(row[col.get("Booked", 7)]).upper() != "TRUE":
                continue
            participants.append({
                "name":   str(row[col.get("User",   0)]).strip() if col.get("User",   0) < len(row) else "",
                "plan":   str(row[col.get("Plan",   4)]).strip() if col.get("Plan",   4) < len(row) else "",
                "volume": str(row[col.get("Volume", 5)]).strip() if col.get("Volume", 5) < len(row) else "",
                "remote": str(row[col.get("Remote", 6)]).lower() == "yes" if col.get("Remote", 6) < len(row) else False,
            })
        result[date_str] = {
            "day_short":   day_info["day_short"],
            "time":        day_info["time"],
            "is_no_train": day_info["is_no_train"],
            "is_remote":   day_info["is_remote"],
            "participants": participants,
        }
    return result

def tr_get_week_data_from_source(week_row: int) -> dict:
    """Читает данные недели прямо из листа 2026 для произвольной строки (напр. прошлая неделя)."""
    days = tr_get_current_week_days(override_row=week_row)
    if not days:
        return {}
    data = get_source_sheet().get_all_values()
    header = data[11] if len(data) > 11 else []
    user_cols = {}
    for name, col_letter in USER_COLUMNS.items():
        idx = ord(col_letter) - ord("A")
        user_cols[name] = idx

    result = {}
    row_idx = week_row
    for day_info in days:
        date_str = day_info["date"]
        if row_idx + 3 >= len(data):
            row_idx += 5
            continue
        bron_row = data[row_idx]
        plan_row = data[row_idx + 1]
        vol_row  = data[row_idx + 2]
        com_row  = data[row_idx + 3]
        participants = []
        for name, col_idx in user_cols.items():
            booked = str(bron_row[col_idx]).upper() == "TRUE" if col_idx < len(bron_row) else False
            if not booked:
                row_idx_next = row_idx
                continue
            plan    = str(plan_row[col_idx]).strip() if col_idx < len(plan_row) else ""
            volume  = str(vol_row[col_idx]).strip()  if col_idx < len(vol_row)  else ""
            comment = str(com_row[col_idx]).strip()  if col_idx < len(com_row)  else ""
            plan    = "" if plan    in ("nan", "None") else plan
            volume  = "" if volume  in ("nan", "None") else volume
            comment = "" if comment in ("nan", "None") else comment
            remote  = day_info["is_remote"] or "удал" in plan.lower() or "удал" in comment.lower()
            participants.append({
                "name": name, "plan": plan, "volume": volume, "remote": remote,
            })
        result[date_str] = {
            "day_short":    day_info["day_short"],
            "time":         day_info["time"],
            "is_no_train":  day_info["is_no_train"],
            "is_remote":    day_info["is_remote"],
            "participants": participants,
        }
        row_idx += 5
    return result

def tr_write_plan_for_date(date_str: str, plan_text: str) -> int:
    """Записывает план всем забронировавшим на указанную дату. Возвращает кол-во записей."""
    rows = find_date_rows(date_str)
    if not rows:
        return 0
    plan_row_1based = rows["bron"] + 1
    _ensure_bd()
    hdr = _bd_rows[1] if len(_bd_rows) >= 2 else []
    col = {h: i for i, h in enumerate(hdr)}
    src = get_source_sheet()
    written = 0
    for row in _bd_rows[2:]:
        if not row or col.get("Date", 1) >= len(row):
            continue
        if row[col["Date"]].strip() != date_str:
            continue
        if str(row[col.get("Booked", 7)]).upper() != "TRUE":
            continue
        user = str(row[col.get("User", 0)]).strip()
        if user in USER_COLUMNS:
            src.update(f"{USER_COLUMNS[user]}{plan_row_1based}", [[plan_text]])
            written += 1
    _invalidate_bd()
    return written

def tr_write_plan_for_user(date_str: str, user_name: str, plan_text: str) -> bool:
    """Записывает план конкретному участнику."""
    rows = find_date_rows(date_str)
    if not rows or user_name not in USER_COLUMNS:
        return False
    plan_row_1based = rows["bron"] + 1
    try:
        get_source_sheet().update(f"{USER_COLUMNS[user_name]}{plan_row_1based}", [[plan_text]])
        _invalidate_bd()
        return True
    except Exception as e:
        print(f"tr_write_plan_for_user({user_name}, {date_str}) ошибка: {e}")
        return False

def tr_write_vol_for_user(date_str: str, user_name: str, vol_text: str) -> bool:
    rows = find_date_rows(date_str)
    if not rows or user_name not in USER_COLUMNS:
        return False
    try:
        get_source_sheet().update(f"{USER_COLUMNS[user_name]}{rows['vol']}", [[vol_text]])
        return True
    except Exception as e:
        print(f"tr_write_vol_for_user({user_name}, {date_str}) ошибка: {e}")
        return False

def tr_write_comm_for_user(date_str: str, user_name: str, comm_text: str) -> bool:
    rows = find_date_rows(date_str)
    if not rows or user_name not in USER_COLUMNS:
        return False
    try:
        get_source_sheet().update(f"{USER_COLUMNS[user_name]}{rows['vol'] + 1}", [[comm_text]])
        return True
    except Exception as e:
        print(f"tr_write_comm_for_user({user_name}, {date_str}) ошибка: {e}")
        return False

def tr_write_volume_for_date(date_str: str, vol_text: str) -> int:
    """Записывает объём всем забронировавшим на указанную дату. Возвращает кол-во записей."""
    rows = find_date_rows(date_str)
    if not rows:
        return 0
    vol_row_1based = rows["vol"]
    _ensure_bd()
    hdr = _bd_rows[1] if len(_bd_rows) >= 2 else []
    col = {h: i for i, h in enumerate(hdr)}
    src = get_source_sheet()
    written = 0
    for row in _bd_rows[2:]:
        if not row or col.get("Date", 1) >= len(row):
            continue
        if row[col["Date"]].strip() != date_str:
            continue
        if str(row[col.get("Booked", 7)]).upper() != "TRUE":
            continue
        user = str(row[col.get("User", 0)]).strip()
        if user in USER_COLUMNS:
            src.update(f"{USER_COLUMNS[user]}{vol_row_1based}", [[vol_text]])
            written += 1
    _invalidate_bd()
    return written

# ═══════════════════════════════════════════════════════════════
# БРОНИРОВАНИЕ В ЛИСТЕ 2026
# ═══════════════════════════════════════════════════════════════

def tr_write_comment_for_date(date_str: str, comment_text: str) -> int:
    """Записывает комментарий всем забронировавшим на указанную дату."""
    rows = find_date_rows(date_str)
    if not rows:
        return 0
    comment_row = rows["vol"] + 1
    _ensure_bd()
    hdr = _bd_rows[1] if len(_bd_rows) >= 2 else []
    col = {h: i for i, h in enumerate(hdr)}
    src = get_source_sheet()
    written = 0
    for row in _bd_rows[2:]:
        if not row or col.get("Date", 1) >= len(row):
            continue
        if row[col["Date"]].strip() != date_str:
            continue
        if str(row[col.get("Booked", 7)]).upper() != "TRUE":
            continue
        user = str(row[col.get("User", 0)]).strip()
        if user in USER_COLUMNS:
            src.update(f"{USER_COLUMNS[user]}{comment_row}", [[comment_text]])
            written += 1
    _invalidate_bd()
    return written

def find_date_rows(date_str: str) -> dict | None:
    try:
        data = get_source_sheet().get_all_values()
    except Exception as e:
        print(f"find_date_rows({date_str}) ошибка: {e}")
        return None
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
# ТРЕНЕР — КЛАВИАТУРЫ
# ═══════════════════════════════════════════════════════════════

def kb_role_select():
    b = InlineKeyboardBuilder()
    b.button(text="🏊 Спортсмен", callback_data="role_student")
    b.button(text="🎓 Тренер",    callback_data="role_trainer")
    b.adjust(2)
    return b.as_markup()

def kb_trainer_menu():
    b = InlineKeyboardBuilder()
    b.button(text="📅 Расписание",   callback_data="tr_schedule")
    b.button(text="👥 Ученики",      callback_data="tr_students")
    b.button(text="🏋️ Тренировки",  callback_data="tr_trainings")
    b.button(text="📊 Статистика",   callback_data="tr_stats")
    b.button(text="📢 Рассылка",     callback_data="tr_broadcast")
    b.adjust(2, 2, 1)
    return b.as_markup()

def kb_persistent_trainer():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🏠 Главное меню"),
                   KeyboardButton(text="👁 Мой вид")]],
        resize_keyboard=True,
    )

def kb_persistent_exit():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🏠 Главное меню"),
                   KeyboardButton(text="🔙 Выйти из просмотра")]],
        resize_keyboard=True,
    )

def kb_tr_week_overview(days: list, week_active: bool, preview: bool = False) -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    if preview:
        for d in days:
            if not d["date"]:
                continue
            short = d["day_short"]
            date  = d["date"]
            if d["is_no_train"]:
                label = f"❌ {short} {date[:5]} — нет тренировки"
            elif d["is_remote"]:
                label = f"🏠 {short} {date[:5]} — удалённо"
            elif d["time"]:
                label = f"🏊 {short} {date[:5]} — ⏰ {d['time']}"
            else:
                label = f"⬜ {short} {date[:5]}"
            b.button(text=label, callback_data=f"tr_day_{date}")
        b.button(text="✅ Открыть эту неделю",   callback_data="tr_confirm_open")
        b.button(text="❌ Свернуть обратно",      callback_data="tr_cancel_expand")
    elif week_active and _week_collapsed:
        b.button(text="📂 Развернуть текущую неделю", callback_data="tr_expand_current")
        b.button(text="⏹ Закрыть неделю",             callback_data="tr_close_week")
        if _prev_week_row != -1:
            b.button(text="📂 Развернуть прошлую неделю", callback_data="tr_view_prev_week")
        b.button(text="📂 Развернуть следующую неделю",   callback_data="tr_expand_next")
        b.button(text="◀️ Назад",                         callback_data="tr_menu")
    else:
        if week_active:
            for d in days:
                if not d["date"]:
                    continue
                short = d["day_short"]
                date  = d["date"]
                if d["is_no_train"]:
                    label = f"❌ {short} {date[:5]} — нет тренировки"
                elif d["is_remote"]:
                    label = f"🏠 {short} {date[:5]} — удалённо"
                elif d["time"]:
                    label = f"🏊 {short} {date[:5]} — ⏰ {d['time']}"
                else:
                    label = f"⬜ {short} {date[:5]}"
                b.button(text=label, callback_data=f"tr_day_{date}")
            b.button(text="⏹ Закрыть неделю", callback_data="tr_close_week")
            b.button(text="📦 Свернуть",        callback_data="tr_collapse_week")
        else:
            if _prev_week_row != -1:
                b.button(text="📂 Развернуть прошлую неделю", callback_data="tr_view_prev_week")
        b.button(text="📂 Развернуть следующую неделю", callback_data="tr_expand_next")
        b.button(text="◀️ Назад",                       callback_data="tr_menu")
    b.adjust(1)
    return b.as_markup()

def kb_tr_prev_week_view(days: list) -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    for d in days:
        if not d["date"]:
            continue
        short = d["day_short"]
        date  = d["date"]
        if d["is_no_train"]:
            label = f"❌ {short} {date[:5]} — нет тренировки"
        elif d["is_remote"]:
            label = f"🏠 {short} {date[:5]} — удалённо"
        elif d["time"]:
            label = f"🏊 {short} {date[:5]} — ⏰ {d['time']}"
        else:
            label = f"⬜ {short} {date[:5]}"
        b.button(text=label, callback_data=f"tr_prevday_{date}")
    b.button(text="📦 Свернуть прошлую неделю", callback_data="tr_collapse_prev")
    b.adjust(1)
    return b.as_markup()

_QUICK_TIMES = [("🕑 14:00", "1400", 14, 0), ("🕗 20:00", "2000", 20, 0)]

def kb_tr_day_edit(date_str: str, has_time: bool) -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    now = _now()
    day_date = _parse_sheet_date(date_str, now.year)
    is_past_day = day_date is not None and day_date < now.date()
    is_today    = day_date is not None and day_date == now.date()

    available = []
    for label, hhmm, hour, minute in _QUICK_TIMES:
        if is_past_day:
            continue  # весь день прошёл — быстрые кнопки не нужны
        if is_today and (now.hour, now.minute) >= (hour, minute):
            continue  # это время уже прошло сегодня
        available.append((label, hhmm))

    for label, hhmm in available:
        b.button(text=label, callback_data=f"tr_qtime_{hhmm}_{date_str}")
    b.button(text="✏️ Другое время",       callback_data=f"tr_custom_time_{date_str}")
    if has_time:
        b.button(text="❌ Нет тренировки", callback_data=f"tr_notraining_{date_str}")
    b.button(text="◀️ Назад к расписанию", callback_data="tr_schedule")
    n = len(available)
    b.adjust(*([2] if n == 2 else [1] * n), 1, *(1 for _ in range(1 + int(has_time))))
    return b.as_markup()

def kb_tr_confirm_action(date_str: str) -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    b.button(text="✅ Подтвердить", callback_data="tr_confirm_time")
    b.button(text="❌ Отменить",    callback_data=f"tr_day_{date_str}")
    b.adjust(2)
    return b.as_markup()

def kb_trainer_students():
    b = InlineKeyboardBuilder()
    for name in USER_COLUMNS:
        grade = _known_grades.get(name, "")
        label = f"{name}  {grade}".strip() if grade else name
        b.button(text=label, callback_data=f"tr_student_{name}")
    b.button(text="◀️ Назад", callback_data="tr_menu")
    b.adjust(2)
    return b.as_markup()

def kb_trainer_student(name: str):
    b = InlineKeyboardBuilder()
    b.button(text="📋 Брони",          callback_data=f"tr_bookings_{name}")
    b.button(text="🎫 Абонемент",       callback_data=f"tr_sub_{name}")
    b.button(text="🏊 Грейд плавание",  callback_data=f"tr_grade_swim_{name}")
    b.button(text="🤿 Грейд DNF/DYN",  callback_data=f"tr_grade_dnf_{name}")
    b.button(text="📊 Состояния",       callback_data=f"tr_results_{name}")
    b.button(text="◀️ Назад",           callback_data="tr_students")
    b.adjust(2, 2, 1, 1)
    return b.as_markup()

def kb_trainer_swim_grades(name: str):
    b = InlineKeyboardBuilder()
    current = _known_grades.get(name, "")
    for i, g in enumerate(SWIM_GRADE_ORDER):
        label = f"✅ {g}" if g == current else g
        b.button(text=label, callback_data=f"tr_setswim_{i}_{name}")
    b.button(text="◀️ Назад", callback_data=f"tr_student_{name}")
    b.adjust(1)
    return b.as_markup()

def kb_trainer_dnf_grades(name: str):
    b = InlineKeyboardBuilder()
    current = _known_dnf_grades.get(name, "")
    for i, g in enumerate(DNF_GRADE_ORDER):
        label = f"✅ {g}" if g == current else g
        b.button(text=label, callback_data=f"tr_setdnf_{i}_{name}")
    b.button(text="🗑 Убрать DNF",  callback_data=f"tr_setdnf_clear_{name}")
    b.button(text="◀️ Назад",       callback_data=f"tr_student_{name}")
    b.adjust(1)
    return b.as_markup()

def kb_trainer_sub(name: str, pending: float = 0):
    b = InlineKeyboardBuilder()
    for delta in (1, 5, 10):
        b.button(text=f"+{delta}", callback_data=f"tr_subdelta_{name}_{delta}")
    if pending:
        b.button(text=f"✅ Сохранить (+{_fmt_sub(pending)})", callback_data=f"tr_subconfirm_{name}")
        b.button(text="↩️ Сбросить", callback_data=f"tr_subcancel_{name}")
        b.button(text="◀️ Назад", callback_data=f"tr_student_{name}")
        b.adjust(3, 1, 1, 1)
    else:
        b.button(text="◀️ Назад", callback_data=f"tr_student_{name}")
        b.adjust(3, 1)
    return b.as_markup()

def kb_trainer_broadcast():
    b = InlineKeyboardBuilder()
    b.button(text="📢 Всем ученикам",    callback_data="tr_broadcast_all")
    b.button(text="👤 Выбрать учеников", callback_data="tr_broadcast_select")
    b.button(text="◀️ Назад",             callback_data="tr_menu")
    b.adjust(1)
    return b.as_markup()

def kb_trainer_broadcast_select(selected: list):
    b = InlineKeyboardBuilder()
    for name in USER_COLUMNS:
        mark = "✅ " if name in selected else ""
        b.button(text=f"{mark}{name}", callback_data=f"tr_bsel_{name}")
    if selected:
        b.button(text=f"📢 Отправить ({len(selected)})", callback_data="tr_bsend")
    b.button(text="◀️ Назад", callback_data="tr_broadcast")
    b.adjust(3)
    return b.as_markup()

# ═══════════════════════════════════════════════════════════════
# БОТ
# ═══════════════════════════════════════════════════════════════

bot = Bot(token=TOKEN)
dp  = Dispatcher()

# ── /start ──────────────────────────────────────────────────────
@dp.message(CommandStart())
async def cmd_start(message: Message):
    uid = message.from_user.id
    if is_trainer(uid):
        await message.answer("—", reply_markup=kb_persistent_trainer())
        await message.answer_photo(photo=FSInputFile("logo.png"),
            caption="🎓 *Панель тренера*", reply_markup=kb_trainer_menu(), parse_mode="Markdown")
        return
    user_name = get_user_name_by_telegram_id(uid)
    if user_name:
        await message.answer("—", reply_markup=kb_persistent())
        await message.answer_photo(photo=FSInputFile("logo.png"),
            caption=f"Привет, {user_name} 👋", reply_markup=kb_main_menu())
    else:
        await message.answer_photo(photo=FSInputFile("logo.png"),
            caption="👋 Добро пожаловать!\n\nКто вы?", reply_markup=kb_role_select())

@dp.message(F.text == "🚀 Начать пользоваться")
async def btn_start(message: Message):
    uid = message.from_user.id
    if is_trainer(uid):
        await message.answer("—", reply_markup=kb_persistent_trainer())
        await message.answer_photo(photo=FSInputFile("logo.png"),
            caption="🎓 *Панель тренера*", reply_markup=kb_trainer_menu(), parse_mode="Markdown")
        return
    user_name = get_user_name_by_telegram_id(uid)
    if user_name:
        await message.answer("—", reply_markup=kb_persistent())
        await message.answer_photo(photo=FSInputFile("logo.png"),
            caption=f"Привет, {user_name} 👋", reply_markup=kb_main_menu())
        return
    await message.answer_photo(photo=FSInputFile("logo.png"),
        caption="Выберите своё имя из списка:", reply_markup=kb_user_list())

@dp.message(F.text == "🏠 Главное меню")
async def btn_main_menu(message: Message):
    uid = message.from_user.id
    if is_trainer(uid) and uid not in _trainer_view_as:
        await message.answer_photo(photo=FSInputFile("logo.png"),
            caption="🎓 *Панель тренера*", reply_markup=kb_trainer_menu(), parse_mode="Markdown")
        return
    user_name = get_user_name_by_telegram_id(uid)
    if user_name:
        await message.answer_photo(photo=FSInputFile("logo.png"),
            caption=f"Привет, {user_name} 👋", reply_markup=kb_main_menu())
    else:
        await message.answer_photo(photo=FSInputFile("logo.png"),
            caption="Выберите своё имя из списка:", reply_markup=kb_user_list())

# ── Выбор роли ──────────────────────────────────────────────────
@dp.callback_query(F.data == "role_student")
async def cb_role_student(callback: CallbackQuery):
    await callback.message.edit_caption(caption="Выберите своё имя из списка:", reply_markup=kb_user_list())
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data == "role_trainer")
async def cb_role_trainer(callback: CallbackQuery):
    uid = callback.from_user.id
    if uid == TRAINER_ID:
        _authenticated_trainers.add(uid)
        await callback.message.edit_caption(caption="🎓 *Панель тренера*",
            reply_markup=kb_trainer_menu(), parse_mode="Markdown")
    elif TRAINER_PASSWORD:
        _trainer_state[uid] = {"action": "waiting_password"}
        b = InlineKeyboardBuilder()
        b.button(text="◀️ Назад", callback_data="tr_back_role")
        await callback.message.edit_caption(caption="🔐 Введите пароль тренера:", reply_markup=b.as_markup())
    else:
        await callback.message.edit_caption(caption="❌ Доступ запрещён.", reply_markup=kb_role_select())
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data == "tr_back_role")
async def cb_tr_back_role(callback: CallbackQuery):
    _trainer_state.pop(callback.from_user.id, None)
    await callback.message.edit_caption(caption="👋 Добро пожаловать!\n\nКто вы?", reply_markup=kb_role_select())
    try: await callback.answer()
    except: pass

# ── Просмотр от имени ученика ────────────────────────────────────
@dp.callback_query(F.data == "tr_view_select")
async def cb_tr_view_select(callback: CallbackQuery):
    uid = callback.from_user.id
    if not is_trainer(uid): return
    # Ищем имя тренера как ученика напрямую (минуя _trainer_view_as)
    name = _user_cache.get(uid)
    if not name:
        try: await callback.answer("Вы не зарегистрированы как ученик", show_alert=True)
        except: pass
        return
    _trainer_view_as[uid] = name
    await callback.message.edit_caption(
        caption=f"👁 *Просмотр как {name}*",
        reply_markup=kb_main_menu(),
        parse_mode="Markdown")
    await callback.message.answer("—", reply_markup=kb_persistent_exit())
    try: await callback.answer()
    except: pass

@dp.message(F.text == "👁 Мой вид")
async def btn_my_view(message: Message):
    uid = message.from_user.id
    if not is_trainer(uid): return
    name = _user_cache.get(uid)
    if not name:
        await message.reply("Вы не зарегистрированы как ученик.")
        return
    _trainer_view_as[uid] = name
    await message.answer("—", reply_markup=kb_persistent_exit())
    await message.answer_photo(
        photo=FSInputFile("logo.png"),
        caption=f"👁 *Просмотр как {name}*",
        reply_markup=kb_main_menu(),
        parse_mode="Markdown")

@dp.message(F.text == "🔙 Выйти из просмотра")
async def btn_exit_view(message: Message):
    uid = message.from_user.id
    _trainer_view_as.pop(uid, None)
    await message.answer("—", reply_markup=kb_persistent_trainer())
    await message.answer_photo(
        photo=FSInputFile("logo.png"),
        caption="🎓 *Панель тренера*",
        reply_markup=kb_trainer_menu(),
        parse_mode="Markdown")

# ── /me и /trainer — переключение вида (команды) ────────────────
@dp.message(Command("me"))
async def cmd_me(message: Message):
    uid = message.from_user.id
    if not is_trainer(uid): return
    name = _user_cache.get(uid)
    if not name:
        await message.reply("Вы не зарегистрированы как ученик.")
        return
    _trainer_view_as[uid] = name
    await message.answer("—", reply_markup=kb_persistent_exit())
    await message.answer_photo(
        photo=FSInputFile("logo.png"),
        caption=f"👁 *Просмотр как {name}*",
        reply_markup=kb_main_menu(),
        parse_mode="Markdown")

@dp.message(Command("trainer"))
async def cmd_trainer(message: Message):
    uid = message.from_user.id
    if not is_trainer(uid): return
    _trainer_view_as.pop(uid, None)
    _trainer_state.pop(uid, None)
    await message.answer("—", reply_markup=kb_persistent_trainer())
    await message.answer_photo(
        photo=FSInputFile("logo.png"),
        caption="🎓 *Панель тренера*",
        reply_markup=kb_trainer_menu(),
        parse_mode="Markdown")

# ── Панель тренера — главное меню ───────────────────────────────
@dp.callback_query(F.data == "tr_menu")
async def cb_tr_menu(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id):
        try: await callback.answer("Доступ запрещён", show_alert=True)
        except: pass
        return
    try:
        week_label = get_bd_sheet().acell("A1").value or "не активна"
    except Exception:
        week_label = "не активна"
    await callback.message.edit_caption(
        caption=f"🎓 *Панель тренера*\n\n📅 {week_label}",
        reply_markup=kb_trainer_menu(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Расписание (обзор недели) ────────────────────────────────────
@dp.callback_query(F.data == "tr_schedule")
async def cb_tr_schedule(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    uid = callback.from_user.id
    _trainer_state.pop(uid, None)   # сбрасываем любое ожидающее состояние
    _ensure_bd()
    week_active = _week_marker_row != -1

    # Режим просмотра прошлой недели
    if _viewing_prev_week and _prev_week_row != -1:
        prev_days = tr_get_current_week_days(override_row=_prev_week_row)
        dates = [d["date"] for d in prev_days if d["date"]]
        dr = f"{dates[0][:5]} — {dates[-1][:5]}" if len(dates) >= 2 else ""
        caption = f"📋 *Прошлая неделя* (просмотр)\n{dr}\n\n📌 Статус: прошлая неделя"
        await callback.message.edit_caption(
            caption=caption,
            reply_markup=kb_tr_prev_week_view(prev_days),
            parse_mode="Markdown")
        try: await callback.answer()
        except: pass
        return

    days = tr_get_current_week_days() if week_active else []

    # Режим предпросмотра: следующая неделя развёрнута, но ещё не открыта
    if _pending_open_row != -1:
        preview_days = tr_get_current_week_days(override_row=_pending_open_row)
        dates = [d["date"] for d in preview_days if d["date"]]
        dr = f"{dates[0][:5]} — {dates[-1][:5]}" if len(dates) >= 2 else ""
        caption = f"📋 *Предпросмотр следующей недели*\n{dr}\n\n📌 Статус: не выбрано\n\nПоставьте времена, затем откройте неделю:"
        await callback.message.edit_caption(
            caption=caption,
            reply_markup=kb_tr_week_overview(preview_days, False, preview=True),
            parse_mode="Markdown")
        try: await callback.answer()
        except: pass
        return

    if week_active and _bd_rows and _bd_rows[0]:
        raw = str(_bd_rows[0][0])
        parts = [p.strip() for p in raw.split("|")]
        week_text = f"{parts[0]} | {parts[1]}" if len(parts) >= 2 else parts[0]
    else:
        week_text = "Неделя не активна"

    status  = "🟢 Активна" if week_active else "🔴 Не активна"
    caption = f"📅 *{week_text}*\n{status}"

    await callback.message.edit_caption(
        caption=caption,
        reply_markup=kb_tr_week_overview(days, week_active),
        parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data == "tr_close_week")
async def cb_tr_close_week(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    ok = tr_close_current_week()
    try: await callback.answer("✅ Неделя закрыта" if ok else "❌ Нет активной недели")
    except: pass
    await cb_tr_schedule(callback)

@dp.callback_query(F.data == "tr_expand_next")
async def cb_tr_expand_next(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    row = tr_expand_next_week()
    if row == -1:
        try: await callback.answer("❌ Следующая неделя не найдена", show_alert=True)
        except: pass
        return
    try: await callback.answer("📂 Неделя развёрнута")
    except: pass
    await cb_tr_schedule(callback)

@dp.callback_query(F.data == "tr_confirm_open")
async def cb_tr_confirm_open(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    ok = tr_confirm_open_week()
    if not ok:
        try: await callback.answer("❌ Ошибка")
        except: pass
        await cb_tr_schedule(callback)
        return
    try: await callback.answer("✅ Неделя открыта!")
    except: pass
    b = InlineKeyboardBuilder()
    b.button(text="📢 Уведомить учеников о расписании", callback_data="tr_notify_students")
    b.button(text="◀️ К расписанию",                    callback_data="tr_schedule")
    b.adjust(1)
    await callback.message.edit_caption(
        caption="✅ *Неделя открыта!*\n\nУчасники ещё не знают — хотите уведомить их?",
        reply_markup=b.as_markup(),
        parse_mode="Markdown")

@dp.callback_query(F.data == "tr_notify_students")
async def cb_tr_notify_students(callback: CallbackQuery):
    global _week_session_notified
    if not is_trainer(callback.from_user.id): return
    _ensure_bd()
    if _bd_rows and _bd_rows[0]:
        raw   = str(_bd_rows[0][0])
        parts = [p.strip() for p in raw.split("|")]
        label = f"{parts[0]} | {parts[1]}" if len(parts) >= 2 else parts[0]
    else:
        label = "новую неделю"
    _week_session_notified = True
    load_all_tids()
    for user_name in USER_COLUMNS:
        await notify_user(
            user_name,
            f"📅 {user_name}, расписание на {label} готово!\n"
            f"Открой бот чтобы посмотреть и записаться 👇"
        )
    try: await callback.answer("📢 Уведомления отправлены!")
    except: pass
    await cb_tr_schedule(callback)

@dp.callback_query(F.data == "tr_cancel_expand")
async def cb_tr_cancel_expand(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    tr_cancel_expand_week()
    try: await callback.answer("↩️ Свёрнуто обратно")
    except: pass
    await cb_tr_schedule(callback)

@dp.callback_query(F.data == "tr_collapse_week")
async def cb_tr_collapse_week(callback: CallbackQuery):
    global _week_collapsed
    if not is_trainer(callback.from_user.id): return
    if _week_marker_row != -1:
        toggle_week_collapse(_week_marker_row, True)
        _week_collapsed = True
        try: await callback.answer("📦 Свёрнуто")
        except: pass
    else:
        try: await callback.answer("❌ Нет активной недели")
        except: pass
    await cb_tr_schedule(callback)

@dp.callback_query(F.data == "tr_expand_current")
async def cb_tr_expand_current(callback: CallbackQuery):
    global _week_collapsed
    if not is_trainer(callback.from_user.id): return
    if _week_marker_row != -1:
        toggle_week_collapse(_week_marker_row, False)
        _week_collapsed = False
        try: await callback.answer("📂 Развёрнуто")
        except: pass
    await cb_tr_schedule(callback)

@dp.callback_query(F.data == "tr_view_prev_week")
async def cb_tr_view_prev_week(callback: CallbackQuery):
    global _viewing_prev_week
    if not is_trainer(callback.from_user.id): return
    if _prev_week_row == -1:
        try: await callback.answer("❌ Нет прошлой недели", show_alert=True)
        except: pass
        return
    toggle_week_collapse(_prev_week_row, False)
    _viewing_prev_week = True
    try: await callback.answer("📂 Прошлая неделя")
    except: pass
    await cb_tr_schedule(callback)

@dp.callback_query(F.data == "tr_collapse_prev")
async def cb_tr_collapse_prev(callback: CallbackQuery):
    global _viewing_prev_week
    if not is_trainer(callback.from_user.id): return
    if _prev_week_row != -1:
        toggle_week_collapse(_prev_week_row, True)
    _viewing_prev_week = False
    try: await callback.answer("📦 Свёрнуто")
    except: pass
    await cb_tr_schedule(callback)

@dp.callback_query(F.data.startswith("tr_prevday_"))
async def cb_tr_prevday(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    date_str = callback.data[len("tr_prevday_"):]
    if _prev_week_row == -1:
        try: await callback.answer()
        except: pass
        return
    days = tr_get_current_week_days(override_row=_prev_week_row)
    d = next((x for x in days if x["date"] == date_str), None)
    if not d:
        try: await callback.answer()
        except: pass
        return
    if d["is_no_train"]:
        info = "❌ нет тренировки"
    elif d["is_remote"]:
        info = "🏠 удалённо"
    elif d["time"]:
        info = f"⏰ {d['time']}"
    else:
        info = "⬜ время не задано"
    try: await callback.answer(f"{d['day_short']} {date_str[:5]}: {info}", show_alert=True)
    except: pass

# ── День — редактирование времени ───────────────────────────────
@dp.callback_query(F.data.startswith("tr_day_"))
async def cb_tr_day(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    uid      = callback.from_user.id
    date_str = callback.data[len("tr_day_"):]
    _trainer_state.pop(uid, None)
    # поддерживаем и режим предпросмотра следующей недели
    days = tr_get_current_week_days(override_row=_pending_open_row) \
           if _pending_open_row != -1 else tr_get_current_week_days()
    d = next((x for x in days if x["date"] == date_str), None)
    if not d:
        try: await callback.answer("❌ День не найден", show_alert=True)
        except: pass
        return
    day_label = f"{d['day_short']} {date_str[:5]}"
    if d["is_no_train"]:
        status = "❌ нет тренировки"
        has_time = True
    elif d["is_remote"]:
        status = "🏠 удалённо"
        has_time = True
    elif d["time"]:
        status = f"⏰ {d['time']}"
        has_time = True
    else:
        status = "⬜ время не задано"
        has_time = False
    await callback.message.edit_caption(
        caption=f"📅 *{day_label}*  —  {status}\n\nВыберите время:",
        reply_markup=kb_tr_day_edit(date_str, has_time),
        parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_qtime_"))
async def cb_tr_qtime(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    # формат: tr_qtime_{hhmm}_{date}
    rest  = callback.data[len("tr_qtime_"):]
    sep   = rest.index("_")
    hhmm, date_str = rest[:sep], rest[sep+1:]
    time_str = f"{hhmm[:2]}:{hhmm[2:]}" if len(hhmm) == 4 else hhmm
    uid  = callback.from_user.id
    _trainer_state[uid] = {"action": "confirm_time", "date": date_str, "time": time_str}
    days = tr_get_current_week_days()
    d = next((x for x in days if x["date"] == date_str), None)
    day_label = f"{d['day_short']} {date_str[:5]}" if d else date_str[:5]
    await callback.message.edit_caption(
        caption=f"⏰ Установить *{time_str}* для {day_label}?",
        reply_markup=kb_tr_confirm_action(date_str),
        parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_custom_time_"))
async def cb_tr_custom_time(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    date_str = callback.data[len("tr_custom_time_"):]
    uid  = callback.from_user.id
    days = tr_get_current_week_days()
    d = next((x for x in days if x["date"] == date_str), None)
    day_label = f"{d['day_short']} {date_str[:5]}" if d else date_str[:5]
    _trainer_state[uid] = {
        "action":     "custom_time",
        "date":       date_str,
        "day_short":  d["day_short"] if d else "",
        "chat_id":    callback.message.chat.id,
        "message_id": callback.message.message_id,
    }
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data=f"tr_day_{date_str}")
    await callback.message.edit_caption(
        caption=f"⏰ *{day_label}* — введите время сообщением\n\nПример: `9:30` или `20:00`",
        reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_notraining_"))
async def cb_tr_notraining(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    date_str = callback.data[len("tr_notraining_"):]
    uid  = callback.from_user.id
    days = tr_get_current_week_days()
    d = next((x for x in days if x["date"] == date_str), None)
    day_label = f"{d['day_short']} {date_str[:5]}" if d else date_str[:5]
    _trainer_state[uid] = {"action": "confirm_time", "date": date_str, "time": "нет тренировки"}
    await callback.message.edit_caption(
        caption=f"❌ Отменить тренировку *{day_label}*?",
        reply_markup=kb_tr_confirm_action(date_str),
        parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data == "tr_confirm_time")
async def cb_tr_confirm_time(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    uid   = callback.from_user.id
    state = _trainer_state.pop(uid, {})
    date_str = state.get("date", "")
    time_str = state.get("time", "")
    if not date_str or time_str == "":
        try: await callback.answer("❌ Нет данных для подтверждения", show_alert=True)
        except: pass
        return
    ok = tr_set_time(date_str, time_str)
    try: await callback.answer("✅ Сохранено" if ok else "❌ Ошибка записи")
    except: pass
    if state.get("from_trainings"):
        await cb_tr_trainings(callback)
    else:
        await cb_tr_schedule(callback)

# ── Ученики ─────────────────────────────────────────────────────
@dp.callback_query(F.data == "tr_students")
async def cb_tr_students(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    await callback.message.edit_caption(
        caption="👥 *Ученики*\n\nВыберите ученика:",
        reply_markup=kb_trainer_students(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_student_") & ~F.data.startswith("tr_students"))
async def cb_tr_student(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    name = callback.data[len("tr_student_"):]
    if name not in USER_COLUMNS:
        try: await callback.answer("Ученик не найден")
        except: pass
        return
    sub  = tr_get_subscription(name)
    swim = _known_grades.get(name, "—")
    dnf  = _known_dnf_grades.get(name, "")
    text = (f"👤 *{name}*\n\n"
            f"🎫 Абонемент: *{sub}* тр.\n"
            f"🏊 Плавание: {swim or '—'}\n"
            + (f"🤿 DNF/DYN: {dnf}\n" if dnf else ""))
    await callback.message.edit_caption(caption=text, reply_markup=kb_trainer_student(name), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_bookings_"))
async def cb_tr_bookings(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    name     = callback.data[len("tr_bookings_"):]
    bookings = tr_get_student_bookings(name)
    if bookings:
        lines = [f"• {r.get('Day','')} {r.get('Date','')} {r.get('Time','')}".strip() for r in bookings]
        text  = f"📋 *Брони {name}:*\n\n" + "\n".join(lines)
    else:
        text = f"📋 *{name}*\n\nЗаписей на эту неделю нет"
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data=f"tr_student_{name}")
    await callback.message.edit_caption(caption=text, reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_sub_"))
async def cb_tr_sub(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    name    = callback.data[len("tr_sub_"):]
    current = tr_get_subscription(name)
    pending = _sub_pending.get(callback.from_user.id, {}).get(name, 0.0)
    caption = f"🎫 *Абонемент — {name}*\n\nОстаток: *{_fmt_sub(current)}* тренировок"
    if pending:
        caption += f"\nДобавляем: *+{_fmt_sub(pending)}* → будет *{_fmt_sub(current + pending)}*"
    await callback.message.edit_caption(
        caption=caption, reply_markup=kb_trainer_sub(name, pending), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_subdelta_"))
async def cb_tr_subdelta(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    rest  = callback.data[len("tr_subdelta_"):]
    delta = int(rest.rsplit("_", 1)[1])
    name  = rest.rsplit("_", 1)[0]
    tid   = callback.from_user.id
    if tid not in _sub_pending:
        _sub_pending[tid] = {}
    _sub_pending[tid][name] = _sub_pending[tid].get(name, 0.0) + delta
    pending = _sub_pending[tid][name]
    current = tr_get_subscription(name)
    caption = (f"🎫 *Абонемент — {name}*\n\n"
               f"Остаток: *{_fmt_sub(current)}* тренировок\n"
               f"Добавляем: *+{_fmt_sub(pending)}* → будет *{_fmt_sub(current + pending)}*")
    await callback.message.edit_caption(
        caption=caption, reply_markup=kb_trainer_sub(name, pending), parse_mode="Markdown")
    try: await callback.answer(f"+{delta}")
    except: pass

@dp.callback_query(F.data.startswith("tr_subconfirm_"))
async def cb_tr_subconfirm(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    name    = callback.data[len("tr_subconfirm_"):]
    tid     = callback.from_user.id
    pending = _sub_pending.get(tid, {}).pop(name, 0.0)
    if tid in _sub_pending and not _sub_pending[tid]:
        del _sub_pending[tid]
    ok = tr_add_subscription(name, pending) if pending else False
    current = tr_get_subscription(name)
    caption = f"🎫 *Абонемент — {name}*\n\nОстаток: *{_fmt_sub(current)}* тренировок"
    if ok:
        caption += f"\n\n✅ Добавлено +{_fmt_sub(pending)}"
        await notify_user(name,
            f"🎫 {name}, ваш абонемент пополнен на {_fmt_sub(pending)} тренировок!\n"
            f"Остаток: *{_fmt_sub(current)}* тренировок.")
    await callback.message.edit_caption(
        caption=caption, reply_markup=kb_trainer_sub(name), parse_mode="Markdown")
    try: await callback.answer("✅ Сохранено" if ok else "Нечего сохранять")
    except: pass

@dp.callback_query(F.data.startswith("tr_subcancel_"))
async def cb_tr_subcancel(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    name = callback.data[len("tr_subcancel_"):]
    tid  = callback.from_user.id
    _sub_pending.get(tid, {}).pop(name, None)
    if tid in _sub_pending and not _sub_pending[tid]:
        del _sub_pending[tid]
    current = tr_get_subscription(name)
    await callback.message.edit_caption(
        caption=f"🎫 *Абонемент — {name}*\n\nОстаток: *{_fmt_sub(current)}* тренировок",
        reply_markup=kb_trainer_sub(name), parse_mode="Markdown")
    try: await callback.answer("Сброшено")
    except: pass

@dp.callback_query(F.data.startswith("tr_grade_swim_"))
async def cb_tr_grade_swim(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    name    = callback.data[len("tr_grade_swim_"):]
    current = _known_grades.get(name, "не задан")
    await callback.message.edit_caption(
        caption=f"🏊 *Грейд плавания — {name}*\n\nТекущий: {current}",
        reply_markup=kb_trainer_swim_grades(name), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_setswim_"))
async def cb_tr_setswim(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    parts = callback.data[len("tr_setswim_"):].split("_", 1)
    idx, name = int(parts[0]), parts[1]
    grade = SWIM_GRADE_ORDER[idx]
    tr_set_grade(name, swim=grade)
    save_grade_history(name, grade)
    await callback.message.edit_caption(
        caption=f"🏊 *Грейд плавания — {name}*\n\nТекущий: {grade}",
        reply_markup=kb_trainer_swim_grades(name), parse_mode="Markdown")
    try: await callback.answer(f"✅ {grade}")
    except: pass

@dp.callback_query(F.data.startswith("tr_grade_dnf_"))
async def cb_tr_grade_dnf(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    name    = callback.data[len("tr_grade_dnf_"):]
    current = _known_dnf_grades.get(name, "не задан")
    await callback.message.edit_caption(
        caption=f"🤿 *Грейд DNF/DYN — {name}*\n\nТекущий: {current}",
        reply_markup=kb_trainer_dnf_grades(name), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_setdnf_"))
async def cb_tr_setdnf(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    rest = callback.data[len("tr_setdnf_"):]
    if rest.startswith("clear_"):
        name  = rest[len("clear_"):]
        grade = ""
    else:
        parts = rest.split("_", 1)
        idx, name = int(parts[0]), parts[1]
        grade = DNF_GRADE_ORDER[idx]
    tr_set_grade(name, dnf=grade)
    if grade:
        save_grade_history(name, grade)
    display = grade or "не задан"
    await callback.message.edit_caption(
        caption=f"🤿 *Грейд DNF/DYN — {name}*\n\nТекущий: {display}",
        reply_markup=kb_trainer_dnf_grades(name), parse_mode="Markdown")
    try: await callback.answer(f"✅ {display}")
    except: pass

@dp.callback_query(F.data.startswith("tr_results_"))
async def cb_tr_results(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    name = callback.data[len("tr_results_"):]
    _ensure_bd()
    lines = []
    hdr = _bd_rows[1] if len(_bd_rows) >= 2 else []
    col = {h: i for i, h in enumerate(hdr)}
    for row in (_bd_rows[2:] if len(_bd_rows) >= 2 else []):
        if not row:
            continue
        r_user   = str(row[col.get("User",   0)]).strip() if col.get("User",   0) < len(row) else ""
        r_booked = str(row[col.get("Booked", 7)]).upper() if col.get("Booked", 7) < len(row) else ""
        if r_user == name and r_booked == "TRUE":
            date = str(row[col.get("Date", 1)]).strip() if col.get("Date", 1) < len(row) else ""
            rows_src = find_date_rows(date)
            state_val = ""
            if rows_src:
                try:
                    state_row = rows_src["bron"] + 4
                    state_val = get_source_sheet().acell(f"{USER_COLUMNS[name]}{state_row}").value or ""
                except Exception:
                    pass
            if state_val and state_val != "-":
                lines.append(f"• {date}: {state_val}")
    text = (f"📊 *Состояния — {name}:*\n\n" + "\n".join(lines[-10:])) if lines \
           else f"📊 *{name}*\n\nОценок пока нет"
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data=f"tr_student_{name}")
    await callback.message.edit_caption(caption=text, reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Рассылка ────────────────────────────────────────────────────
@dp.callback_query(F.data == "tr_broadcast")
async def cb_tr_broadcast(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    await callback.message.edit_caption(
        caption="📢 *Рассылка*\n\nВыберите получателей:",
        reply_markup=kb_trainer_broadcast(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data == "tr_broadcast_all")
async def cb_tr_broadcast_all(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    _trainer_state[callback.from_user.id] = {"action": "broadcast", "recipients": list(USER_COLUMNS.keys())}
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Отмена", callback_data="tr_broadcast")
    await callback.message.edit_caption(
        caption=f"📢 *Рассылка всем ({len(USER_COLUMNS)})*\n\nВведите текст сообщения:",
        reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data == "tr_broadcast_select")
async def cb_tr_broadcast_select(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    state = _trainer_state.get(callback.from_user.id, {})
    if state.get("action") != "broadcast_select":
        _trainer_state[callback.from_user.id] = {"action": "broadcast_select", "selected": []}
    selected = _trainer_state[callback.from_user.id].get("selected", [])
    await callback.message.edit_caption(
        caption=f"👤 Выберите учеников ({len(selected)} выбрано):",
        reply_markup=kb_trainer_broadcast_select(selected))
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data.startswith("tr_bsel_"))
async def cb_tr_bsel(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    name  = callback.data[len("tr_bsel_"):]
    state = _trainer_state.setdefault(callback.from_user.id, {"action": "broadcast_select", "selected": []})
    selected = state.setdefault("selected", [])
    if name in selected:
        selected.remove(name)
    else:
        selected.append(name)
    await callback.message.edit_caption(
        caption=f"👤 Выберите учеников ({len(selected)} выбрано):",
        reply_markup=kb_trainer_broadcast_select(selected))
    try: await callback.answer()
    except: pass

@dp.callback_query(F.data == "tr_bsend")
async def cb_tr_bsend(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    state    = _trainer_state.get(callback.from_user.id, {})
    selected = state.get("selected", [])
    if not selected:
        try: await callback.answer("Никого не выбрано", show_alert=True)
        except: pass
        return
    _trainer_state[callback.from_user.id] = {"action": "broadcast", "recipients": selected}
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Отмена", callback_data="tr_broadcast_select")
    await callback.message.edit_caption(
        caption=f"📢 Рассылка *{len(selected)}* ученикам\n\nВведите текст:",
        reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Статистика ──────────────────────────────────────────────────
@dp.callback_query(F.data == "tr_stats")
async def cb_tr_stats(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    stats = tr_get_week_stats()
    if stats:
        lines = [f"• {n}: {c} тр." for n, c in sorted(stats.items(), key=lambda x: -x[1])]
        text  = (f"📊 *Статистика недели*\n\n"
                 f"Всего записей: {sum(stats.values())}\n"
                 f"Активных: {len(stats)}/{len(USER_COLUMNS)}\n\n"
                 + "\n".join(lines))
    else:
        text = "📊 *Статистика*\n\nДанных нет"
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data="tr_menu")
    await callback.message.edit_caption(caption=text, reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Тренировки — обзор недели ────────────────────────────────────
@dp.callback_query(F.data == "tr_trainings")
async def cb_tr_trainings(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    _ensure_bd()

    uid = callback.from_user.id

    if not _bd_rows or not _bd_rows[0]:
        b = InlineKeyboardBuilder()
        b.button(text="◀️ Назад", callback_data="tr_menu")
        b.adjust(1)
        await callback.message.edit_caption(
            caption="🏋️ <b>Тренировки</b>\n\n⏳ Расписание ещё не готово.",
            reply_markup=b.as_markup(), parse_mode="HTML")
        try: await callback.answer()
        except: pass
        return

    viewing_prev = uid in _tr_viewing_prev and _prev_week_row != -1
    if viewing_prev:
        week_data = tr_get_week_data_from_source(_prev_week_row)
        week_text = "📋 <b>Прошлая неделя</b>"
    else:
        week_data = tr_get_training_week_data()
        raw = str(_bd_rows[0][0])
        parts = [p.strip() for p in raw.split("|")]
        week_text = html.escape(f"{parts[0]} | {parts[1]}" if len(parts) >= 2 else parts[0])

    total_booked = sum(len(d["participants"]) for d in week_data.values())
    missing = []
    for date_str, d in week_data.items():
        for p in d["participants"]:
            if not p["plan"]:
                missing.append(f"{p['name']} ({d['day_short'].lower()})")

    caption = f"🏋️ <b>Тренировки</b>\n\n{week_text}\n\n"
    caption += f"На этой неделе: <b>{total_booked}</b> тренировок\n"
    if missing:
        miss_str = html.escape(", ".join(missing[:6]))
        if len(missing) > 6:
            miss_str += f" +{len(missing) - 6}"
        caption += f"⚠️ Нет плана: {miss_str}"
    elif total_booked > 0:
        caption += "✅ У всех написан план"

    rows_kb = []
    week_active = _week_marker_row != -1
    expanded_date = _tr_expanded.get(uid)

    # Навигация — прошлая неделя сверху
    if viewing_prev:
        rows_kb.append([InlineKeyboardButton(text="📅 Текущая неделя", callback_data="tr_tcurr")])
    elif _prev_week_row != -1:
        rows_kb.append([InlineKeyboardButton(text="📂 Прошлая неделя", callback_data="tr_tprev")])

    days_ordered = tr_get_current_week_days() if not viewing_prev else tr_get_current_week_days(override_row=_prev_week_row)
    for day_info in days_ordered:
        date_str = day_info["date"]
        if not date_str:
            continue
        d = week_data.get(date_str)
        if not d:
            continue
        short = d["day_short"]

        if d["is_no_train"] and not d["participants"]:
            rows_kb.append([InlineKeyboardButton(
                text=f"❌ {short} {date_str[:5]} — нет тренировки",
                callback_data="noop",
            )])
            continue

        # Заголовок дня — кликабельный аккордеон
        time_part = d["time"] if d["time"] and not d["is_no_train"] else (
            "удалённо" if d["is_remote"] or d["participants"] else ""
        )
        n = len(d["participants"])
        n_str = f" — {n} уч." if n else ""
        icon = "🏠" if d["is_remote"] else "🏊"
        is_expanded = date_str == expanded_date
        exp_icon = "🔽" if is_expanded else "▶️"
        rows_kb.append([InlineKeyboardButton(
            text=f"{exp_icon} {icon} {short} {date_str[:5]}{time_part}{n_str}",
            callback_data=f"tr_texpand_{date_str}",
        )])

        if is_expanded and d["participants"]:
            # Кнопки участников (до 3 в ряд)
            # С планом → карточка; без плана → чекбокс для мультивыбора
            sel_for_day = _tr_plan_sel.get(uid, {}).get(date_str, [])
            part_row = []
            for p in d["participants"]:
                remote_mark = " 🏠" if p["remote"] else ""
                if p["plan"]:
                    part_row.append(InlineKeyboardButton(
                        text=f"✅ {p['name']}{remote_mark}",
                        callback_data=f"tr_tpart_{date_str}_{p['name']}",
                    ))
                else:
                    mark = "☑️" if p["name"] in sel_for_day else "◻️"
                    part_row.append(InlineKeyboardButton(
                        text=f"{mark} {p['name']}{remote_mark}",
                        callback_data=f"tr_ptog_{date_str}_{p['name']}",
                    ))
                if len(part_row) == 3:
                    rows_kb.append(part_row)
                    part_row = []
            if part_row:
                rows_kb.append(part_row)
            if sel_for_day:
                rows_kb.append([InlineKeyboardButton(
                    text=f"✏️ Написать план ({len(sel_for_day)} уч.)",
                    callback_data=f"tr_pwrite_{date_str}",
                )])

    if week_active:
        if _week_collapsed:
            rows_kb.append([InlineKeyboardButton(text="📂 Развернуть текущую неделю", callback_data="tr_tcurr_expand")])
        else:
            rows_kb.append([InlineKeyboardButton(text="📦 Свернуть текущую неделю", callback_data="tr_tcollapse")])
    rows_kb.append([InlineKeyboardButton(text="◀️ Назад", callback_data="tr_menu")])

    try:
        await callback.message.edit_caption(
            caption=caption,
            reply_markup=InlineKeyboardMarkup(inline_keyboard=rows_kb),
            parse_mode="HTML",
        )
    except Exception as e:
        print(f"cb_tr_trainings edit_caption error: {e}")
    try: await callback.answer()
    except: pass

# ── Тренировки — раскрыть/свернуть день (аккордеон) ─────────────
@dp.callback_query(F.data.startswith("tr_texpand_"))
async def cb_tr_texpand(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    uid = callback.from_user.id
    date_str = callback.data[len("tr_texpand_"):]
    if _tr_expanded.get(uid) == date_str:
        _tr_expanded.pop(uid, None)
        _tr_plan_sel.get(uid, {}).pop(date_str, None)  # сброс выбора при сворачивании
    else:
        _tr_expanded[uid] = date_str
    await cb_tr_trainings(callback)

# ── Тренировки — переключить участника в мультивыборе плана ──────
@dp.callback_query(F.data.startswith("tr_ptog_"))
async def cb_tr_ptog(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    uid      = callback.from_user.id
    rest     = callback.data[len("tr_ptog_"):]
    date_str = rest[:10]
    name     = rest[11:]
    day_sel  = _tr_plan_sel.setdefault(uid, {}).setdefault(date_str, [])
    if name in day_sel:
        day_sel.remove(name)
    else:
        day_sel.append(name)
    await cb_tr_trainings(callback)

# ── Тренировки — написать план для выбранных участников ──────────
@dp.callback_query(F.data.startswith("tr_pwrite_"))
async def cb_tr_pwrite(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    uid      = callback.from_user.id
    date_str = callback.data[len("tr_pwrite_"):]
    selected = _tr_plan_sel.get(uid, {}).get(date_str, [])
    if not selected:
        try: await callback.answer("❌ Никто не выбран", show_alert=True)
        except: pass
        return
    _ensure_bd()
    week_data = tr_get_training_week_data()
    d         = week_data.get(date_str, {})
    day_label = f"{d.get('day_short', '')} {date_str[:5]}"
    names_str = ", ".join(selected)
    _trainer_state[uid] = {
        "action":     "write_plan",
        "date":       date_str,
        "names":      list(selected),
        "chat_id":    callback.message.chat.id,
        "message_id": callback.message.message_id,
    }
    hint = "После `---` на отдельной строке — объём (запишется автоматически)."
    caption = (f"✏️ *Написать план — {day_label}*\n\n"
               f"Для: _{names_str}_\n\n{hint}\n\nВведите текст плана:")
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data="tr_trainings")
    await callback.message.edit_caption(caption=caption, reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Тренировки — карточка дня ────────────────────────────────────
@dp.callback_query(F.data.startswith("tr_tday_"))
async def cb_tr_tday(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    date_str = callback.data[len("tr_tday_"):]
    _ensure_bd()
    week_data = tr_get_training_week_data()
    d = week_data.get(date_str)
    if not d:
        try: await callback.answer("День не найден", show_alert=True)
        except: pass
        return

    short = d["day_short"]
    time_label = d["time"] if d["time"] else ("удалённо" if d["is_remote"] else "")
    header = f"🏋️ *{short} {date_str[:5]}*" + (f" — {time_label}" if time_label else "")
    caption = f"{header}\n\n"

    if not d["participants"]:
        caption += "👥 Никто не записан\n"
    else:
        caption += "👥 *Участники:*\n"
        for p in d["participants"]:
            plan_icon = "✅" if p["plan"] else "⚠️"
            remote_tag = " 🏠" if p["remote"] else ""
            caption += f"• {p['name']}{remote_tag} {plan_icon}\n"
        plans = [p["plan"] for p in d["participants"] if p["plan"]]
        if plans:
            caption += f"\n📋 *План:*\n{plans[0]}\n"
        else:
            caption += "\n📋 *План:* не написан\n"
        vols = [p["volume"] for p in d["participants"] if p["volume"]]
        if vols:
            caption += f"\n📏 *Объём:* {vols[0]}\n"

    b = InlineKeyboardBuilder()
    if d["participants"]:
        has_plan = any(p["plan"] for p in d["participants"])
        b.button(
            text="✏️ Редактировать план" if has_plan else "✏️ Написать план",
            callback_data=f"tr_wplan_{date_str}",
        )
        b.button(text="📏 Задать объём", callback_data=f"tr_wvol_{date_str}")
    if not d["is_no_train"] and d["time"]:
        b.button(text="❌ Нет тренировки", callback_data=f"tr_tnoday_{date_str}")
    b.button(text="◀️ Назад к тренировкам", callback_data="tr_trainings")
    b.adjust(1)

    await callback.message.edit_caption(caption=caption, reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Тренировки — карточка отдельного участника ───────────────────
@dp.callback_query(F.data.startswith("tr_tpart_"))
async def cb_tr_tpart(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    rest      = callback.data[len("tr_tpart_"):]
    date_str  = rest[:10]   # DD.MM.YYYY — всегда 10 символов
    name      = rest[11:]   # после разделителя _
    _ensure_bd()
    week_data = tr_get_training_week_data()
    d = week_data.get(date_str)
    if not d:
        try: await callback.answer("День не найден", show_alert=True)
        except: pass
        return
    p = next((x for x in d["participants"] if x["name"] == name), None)
    if not p:
        try: await callback.answer("Участник не найден", show_alert=True)
        except: pass
        return

    short = d["day_short"]
    time_label = d["time"] if d["time"] else "удалённо"
    remote_tag = " 🏠 удалённо" if p["remote"] else ""
    caption = f"👤 *{name}* — {short} {date_str[:5]} {time_label}{remote_tag}\n\n"
    if p["plan"]:
        caption += f"📋 *План:*\n{p['plan']}\n\n"
    else:
        caption += "📋 *План:* не написан\n\n"
    if p["volume"]:
        caption += f"📏 *Объём:* {p['volume']}\n"

    b = InlineKeyboardBuilder()
    plan_btn = "✏️ Редактировать план" if p["plan"] else "📋 Написать план"
    b.button(text=plan_btn, callback_data=f"tr_wplan_u_{date_str}_{name}")
    b.button(text="◀️ Назад к тренировкам", callback_data="tr_trainings")
    b.adjust(1)
    await callback.message.edit_caption(caption=caption, reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Тренировки — написать план конкретному участнику ─────────────
@dp.callback_query(F.data.startswith("tr_wplan_u_"))
async def cb_tr_wplan_user(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    rest     = callback.data[len("tr_wplan_u_"):]
    date_str = rest[:10]
    name     = rest[11:]
    uid = callback.from_user.id
    _ensure_bd()
    week_data = tr_get_training_week_data()
    d = week_data.get(date_str, {})
    day_label = f"{d.get('day_short', '')} {date_str[:5]}"
    p = next((x for x in d.get("participants", []) if x["name"] == name), None)
    current = p["plan"] if p else ""

    _trainer_state[uid] = {
        "action":     "write_plan",
        "date":       date_str,
        "name":       name,
        "chat_id":    callback.message.chat.id,
        "message_id": callback.message.message_id,
    }
    b = InlineKeyboardBuilder()
    hint = "После `---` на отдельной строке — объём (запишется автоматически)."
    if current:
        caption = (f"✏️ *Редактировать план — {name}, {day_label}*\n\n"
                   f"Текущий:\n_{current}_\n\n{hint}\n\nВведите новый текст:")
    else:
        caption = (f"✏️ *Написать план — {name}, {day_label}*\n\n"
                   f"{hint}\n\nВведите текст плана:")
    b.button(text="◀️ Назад", callback_data=f"tr_tpart_{date_str}_{name}")
    await callback.message.edit_caption(caption=caption, reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Тренировки — написать/редактировать план ─────────────────────
@dp.callback_query(F.data.startswith("tr_wplan_"))
async def cb_tr_wplan(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    date_str = callback.data[len("tr_wplan_"):]
    uid = callback.from_user.id
    _ensure_bd()
    week_data = tr_get_training_week_data()
    d = week_data.get(date_str, {})
    day_label = f"{d.get('day_short', '')} {date_str[:5]}"
    plans = [p["plan"] for p in d.get("participants", []) if p["plan"]]
    current = plans[0] if plans else ""

    _trainer_state[uid] = {
        "action":     "write_plan",
        "date":       date_str,
        "chat_id":    callback.message.chat.id,
        "message_id": callback.message.message_id,
    }
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data="tr_trainings")
    hint = "После `---` на отдельной строке — объём (запишется автоматически)."
    if current:
        caption = (f"✏️ *Редактировать план — {day_label}*\n\n"
                   f"Текущий:\n_{current}_\n\n{hint}\n\nВведите новый текст:")
    else:
        caption = (f"✏️ *Написать план — {day_label}*\n\n"
                   f"{hint}\n\nВведите текст плана:")
    await callback.message.edit_caption(caption=caption, reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Тренировки — задать объём ────────────────────────────────────
@dp.callback_query(F.data.startswith("tr_wvol_"))
async def cb_tr_wvol(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    date_str = callback.data[len("tr_wvol_"):]
    uid = callback.from_user.id
    _ensure_bd()
    week_data = tr_get_training_week_data()
    d = week_data.get(date_str, {})
    day_label = f"{d.get('day_short', '')} {date_str[:5]}"
    vols = [p["volume"] for p in d.get("participants", []) if p["volume"]]
    current = vols[0] if vols else ""

    _trainer_state[uid] = {
        "action":     "write_volume",
        "date":       date_str,
        "chat_id":    callback.message.chat.id,
        "message_id": callback.message.message_id,
    }
    b = InlineKeyboardBuilder()
    b.button(text="◀️ Назад", callback_data="tr_trainings")
    if current:
        caption = (f"📏 *Редактировать объём — {day_label}*\n\n"
                   f"Текущий: _{current}_\n\nВведите новый объём:")
    else:
        caption = (f"📏 *Задать объём — {day_label}*\n\n"
                   f"Введите объём (пример: `3000 м`):")
    await callback.message.edit_caption(caption=caption, reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Тренировки — подтвердить план ───────────────────────────────
@dp.callback_query(F.data == "tr_cplan")
async def cb_tr_cplan(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    uid = callback.from_user.id
    state = _trainer_state.pop(uid, {})
    date_str  = state.get("date", "")
    plan_text = state.get("text", "")
    if not date_str or not plan_text:
        try: await callback.answer("❌ Нет данных")
        except: pass
        return
    name      = state.get("name", "")
    names     = state.get("names", [])
    vol_text  = state.get("vol", "")
    comm_text = state.get("comm", "")
    if name:
        ok = tr_write_plan_for_user(date_str, name, plan_text)
        if vol_text:
            tr_write_vol_for_user(date_str, name, vol_text)
        if comm_text:
            tr_write_comm_for_user(date_str, name, comm_text)
        try: await callback.answer("✅ План сохранён" if ok else "❌ Ошибка записи")
        except: pass
    elif names:
        written = sum(1 for n in names if tr_write_plan_for_user(date_str, n, plan_text))
        if vol_text:
            for n in names:
                tr_write_vol_for_user(date_str, n, vol_text)
        if comm_text:
            for n in names:
                tr_write_comm_for_user(date_str, n, comm_text)
        _tr_plan_sel.get(uid, {}).pop(date_str, None)
        try: await callback.answer(f"✅ План записан для {written} уч.")
        except: pass
    else:
        count = tr_write_plan_for_date(date_str, plan_text)
        if vol_text:
            tr_write_volume_for_date(date_str, vol_text)
        if comm_text:
            tr_write_comment_for_date(date_str, comm_text)
        try: await callback.answer(f"✅ Сохранено для {count} уч." if count else "❌ Ошибка записи")
        except: pass
    _tr_expanded[uid] = date_str
    nav_chat_id = state.get("chat_id") or (callback.message.chat.id if callback.message else None)
    nav_msg_id  = state.get("message_id") or (callback.message.message_id if callback.message else None)
    b_nav = InlineKeyboardBuilder()
    b_nav.button(text="🏋️ К тренировкам", callback_data="tr_trainings")
    nav_shown = False
    if nav_chat_id and nav_msg_id:
        try:
            await bot.edit_message_caption(
                chat_id=nav_chat_id, message_id=nav_msg_id,
                caption="✅ <b>Сохранено!</b>",
                reply_markup=b_nav.as_markup(), parse_mode="HTML"
            )
            nav_shown = True
        except Exception as e:
            print(f"cb_tr_cplan nav edit error: {e}")
    if not nav_shown:
        try:
            await bot.send_message(
                callback.from_user.id, "✅ <b>Сохранено!</b>",
                reply_markup=b_nav.as_markup(), parse_mode="HTML"
            )
        except Exception as e:
            print(f"cb_tr_cplan send_message error: {e}")

# ── Тренировки — подтвердить объём ──────────────────────────────
@dp.callback_query(F.data == "tr_cvol")
async def cb_tr_cvol(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    uid = callback.from_user.id
    state = _trainer_state.pop(uid, {})
    date_str = state.get("date", "")
    vol_text  = state.get("text", "")
    if not date_str or not vol_text:
        try: await callback.answer("❌ Нет данных")
        except: pass
        return
    count = tr_write_volume_for_date(date_str, vol_text)
    try: await callback.answer(f"✅ Сохранено для {count} уч." if count else "❌ Ошибка записи")
    except: pass
    callback.data = "tr_trainings"
    await cb_tr_trainings(callback)

# ── Тренировки — отменить тренировку в день ─────────────────────
@dp.callback_query(F.data.startswith("tr_tnoday_"))
async def cb_tr_tnoday(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    date_str = callback.data[len("tr_tnoday_"):]
    uid = callback.from_user.id
    _ensure_bd()
    week_data = tr_get_training_week_data()
    d = week_data.get(date_str, {})
    day_label = f"{d.get('day_short', '')} {date_str[:5]}"

    _trainer_state[uid] = {
        "action": "confirm_time",
        "date": date_str,
        "time": "нет тренировки",
        "from_trainings": True,
    }
    b = InlineKeyboardBuilder()
    b.button(text="✅ Подтвердить", callback_data="tr_confirm_time")
    b.button(text="❌ Отменить",    callback_data="tr_trainings")
    b.adjust(2)
    await callback.message.edit_caption(
        caption=f"❌ Отменить тренировку *{day_label}*?",
        reply_markup=b.as_markup(), parse_mode="Markdown")
    try: await callback.answer()
    except: pass

# ── Тренировки — навигация по неделям ───────────────────────────
@dp.callback_query(F.data == "tr_tcollapse")
async def cb_tr_tcollapse(callback: CallbackQuery):
    global _week_collapsed
    if not is_trainer(callback.from_user.id): return
    if _week_marker_row != -1:
        toggle_week_collapse(_week_marker_row, True)
        _week_collapsed = True
        try: await callback.answer("📦 Свёрнуто")
        except: pass
    await cb_tr_trainings(callback)

@dp.callback_query(F.data == "tr_tcurr_expand")
async def cb_tr_texpand_curr(callback: CallbackQuery):
    global _week_collapsed
    if not is_trainer(callback.from_user.id): return
    if _week_marker_row != -1:
        toggle_week_collapse(_week_marker_row, False)
        _week_collapsed = False
        try: await callback.answer("📂 Развёрнуто")
        except: pass
    await cb_tr_trainings(callback)

@dp.callback_query(F.data == "tr_tprev")
async def cb_tr_tprev(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    if _prev_week_row == -1:
        try: await callback.answer("❌ Нет прошлой недели", show_alert=True)
        except: pass
        return
    toggle_week_collapse(_prev_week_row, False)
    _tr_viewing_prev.add(callback.from_user.id)
    _tr_expanded.pop(callback.from_user.id, None)
    try: await callback.answer("📂 Прошлая неделя")
    except: pass
    await cb_tr_trainings(callback)

@dp.callback_query(F.data == "tr_tcurr")
async def cb_tr_tcurr(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    _tr_viewing_prev.discard(callback.from_user.id)
    _tr_expanded.pop(callback.from_user.id, None)
    try: await callback.answer("📅 Текущая неделя")
    except: pass
    await cb_tr_trainings(callback)

@dp.callback_query(F.data == "tr_tnext")
async def cb_tr_tnext(callback: CallbackQuery):
    if not is_trainer(callback.from_user.id): return
    row = tr_expand_next_week()
    if row == -1:
        try: await callback.answer("❌ Следующая неделя не найдена", show_alert=True)
        except: pass
    else:
        try: await callback.answer("📂 Следующая неделя развёрнута")
        except: pass
    await cb_tr_trainings(callback)

# ── Текстовые сообщения тренера (пароль, рассылка, время) ───────
@dp.message(lambda m: m.from_user.id in _trainer_state)
async def handle_trainer_input(message: Message):
    uid   = message.from_user.id
    state = _trainer_state.get(uid, {})
    action = state.get("action", "")

    if action == "waiting_password":
        if message.text and message.text.strip() == TRAINER_PASSWORD:
            _authenticated_trainers.add(uid)
            _trainer_state.pop(uid, None)
            await message.answer("—", reply_markup=kb_persistent())
            await message.answer_photo(photo=FSInputFile("logo.png"),
                caption="✅ *Добро пожаловать в панель тренера!*",
                reply_markup=kb_trainer_menu(), parse_mode="Markdown")
        else:
            await message.reply("❌ Неверный пароль. Попробуйте ещё раз:")

    elif action == "broadcast":
        text       = message.text or ""
        recipients = state.get("recipients", [])
        _trainer_state.pop(uid, None)
        sent = 0
        for name in recipients:
            try:
                await notify_user(name, f"📢 *Сообщение от тренера:*\n\n{text}", parse_mode="Markdown")
                sent += 1
            except Exception:
                pass
        await message.reply(f"✅ Отправлено {sent}/{len(recipients)} ученикам")

    elif action == "set_time":
        date_str = state.get("date", "")
        time_str = (message.text or "").strip()
        _trainer_state.pop(uid, None)
        ok = tr_set_time(date_str, time_str)
        if ok:
            await message.reply(f"✅ Время для *{date_str}* установлено: `{time_str}`", parse_mode="Markdown")
        else:
            await message.reply(f"❌ Не удалось обновить время для {date_str}")

    elif action == "custom_time":
        date_str  = state.get("date", "")
        day_short = state.get("day_short", "")
        chat_id   = state.get("chat_id")
        msg_id    = state.get("message_id")
        raw = (message.text or "").strip()
        # нормализуем "9 30" → "9:30"
        normalized = re.sub(r"^(\d{1,2})\s+(\d{2})$", r"\1:\2", raw)
        if not re.match(r"^\d{1,2}:\d{2}$", normalized):
            await message.reply("❌ Неверный формат. Пример: `9:30` или `20:00`", parse_mode="Markdown")
            return
        _trainer_state[uid] = {"action": "confirm_time", "date": date_str, "time": normalized}
        day_label = f"{day_short} {date_str[:5]}" if day_short else date_str[:5]
        try:
            await bot.edit_message_caption(
                chat_id=chat_id, message_id=msg_id,
                caption=f"⏰ Установить *{normalized}* для {day_label}?",
                reply_markup=kb_tr_confirm_action(date_str),
                parse_mode="Markdown")
        except Exception as e:
            print(f"handle_trainer_input custom_time edit error: {e}")

    elif action == "write_plan":
        date_str  = state.get("date", "")
        raw_text  = (message.text or "").strip()
        chat_id   = state.get("chat_id")
        msg_id    = state.get("message_id")
        if not raw_text:
            await message.reply("❌ Текст не может быть пустым.")
            return
        parts = re.split(r'\n?(?:---+|—-)\n?', raw_text, maxsplit=2)
        plan_text = parts[0].strip()
        vol_text  = parts[1].strip() if len(parts) > 1 else ""
        comm_text = parts[2].strip() if len(parts) > 2 else ""
        _trainer_state[uid] = {
            "action":     "confirm_plan",
            "date":       date_str,
            "text":       plan_text,
            "vol":        vol_text,
            "comm":       comm_text,
            "name":       state.get("name", ""),
            "names":      state.get("names", []),
            "chat_id":    chat_id,
            "message_id": msg_id,
        }
        preview = f"📋 <b>План:</b>\n{html.escape(plan_text)}"
        if vol_text:
            preview += f"\n\n📏 <b>Объём:</b> {html.escape(vol_text)}"
        if comm_text:
            preview += f"\n\n💬 <b>Комментарий:</b> {html.escape(comm_text)}"
        b = InlineKeyboardBuilder()
        b.button(text="✅ Принять",  callback_data="tr_cplan")
        b.button(text="❌ Отменить", callback_data="tr_trainings")
        b.adjust(2)
        try:
            await bot.edit_message_caption(
                chat_id=chat_id, message_id=msg_id,
                caption=f"✏️ <b>Сохранить?</b>\n\n{preview}",
                reply_markup=b.as_markup(), parse_mode="HTML")
        except Exception as e:
            print(f"handle_trainer_input write_plan error: {e}")
            await message.reply(f"❌ Не удалось показать превью: {e}")

    elif action == "write_volume":
        date_str = state.get("date", "")
        vol_text = (message.text or "").strip()
        chat_id  = state.get("chat_id")
        msg_id   = state.get("message_id")
        if not vol_text:
            await message.reply("❌ Текст не может быть пустым.")
            return
        _trainer_state[uid] = {"action": "confirm_volume", "date": date_str, "text": vol_text}
        b = InlineKeyboardBuilder()
        b.button(text="✅ Принять",  callback_data="tr_cvol")
        b.button(text="❌ Отменить", callback_data="tr_trainings")
        b.adjust(2)
        try:
            await bot.edit_message_caption(
                chat_id=chat_id, message_id=msg_id,
                caption=f"📏 <b>Сохранить объём?</b>\n\n{html.escape(vol_text)}",
                reply_markup=b.as_markup(), parse_mode="HTML")
        except Exception as e:
            print(f"handle_trainer_input write_volume error: {e}")
            await message.reply(f"❌ Не удалось показать превью: {e}")

    elif action in ("confirm_plan", "confirm_volume"):
        await message.reply("Нажмите ✅ *Принять* или ❌ *Отменить* в сообщении выше.", parse_mode="Markdown")

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
        if t and (not t["time"] or t["remote"]):
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
        col       = _sub_col(user_name)
        remaining = get_sub_sheet().acell(f"{col}5").value or "0"
        left      = float(str(remaining).replace(" ", "").replace(",", "."))
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
                    if key not in _known_plans:
                        _known_plans[key] = new_plan  # baseline для новой даты — без уведомления
                        continue
                    old_plan = _known_plans[key]
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
    """Каждые 5 минут проверяет изменение грейдов в листе GradeCurrent."""
    while True:
        await asyncio.sleep(300)
        try:
            ws = get_grade_current_sheet()
            rows = ws.get_all_values()
            for row in rows[1:]:
                if not row or not row[0].strip():
                    continue
                name     = row[0].strip()
                new_swim = _extract_swim_grade(row[1].strip() if len(row) > 1 else "")
                new_dnf  = _extract_dnf_grade(row[2].strip() if len(row) > 2 else "")
                old_swim = _known_grades.get(name)
                old_dnf  = _known_dnf_grades.get(name)

                if old_swim is None:
                    _known_grades[name] = new_swim
                    if new_swim:
                        save_grade_history(name, new_swim)
                elif new_swim != old_swim:
                    _known_grades[name] = new_swim
                    if new_swim:
                        save_grade_history(name, new_swim)
                        await notify_user(
                            name,
                            f"🏅 Поздравляем, {name}!\n\nУ вас новый уровень: *{new_swim}*",
                            parse_mode="Markdown",
                        )

                if old_dnf is None:
                    _known_dnf_grades[name] = new_dnf
                    if new_dnf:
                        save_grade_history(name, new_dnf)
                elif new_dnf != old_dnf:
                    _known_dnf_grades[name] = new_dnf
                    if new_dnf:
                        save_grade_history(name, new_dnf)
                        await notify_user(
                            name,
                            f"🏅 Поздравляем, {name}!\n\nУ вас новый уровень: *{new_dnf}*",
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
    load_week_rows()
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
