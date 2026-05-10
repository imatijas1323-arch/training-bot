"""
sync_bd.py
----------
Читает лист "2026" из Google Таблицы, находит текущую неделю,
формирует плоскую таблицу (User, Date, Day, Time, Plan, Volume, Remote, Booked, Comments)
и записывает её в лист "bd".
 
Зависимости:
    pip install gspread google-auth
 
Настройка сервисного аккаунта:
    1. Google Cloud Console → APIs & Services → Credentials → Create credentials → Service account
    2. Скачать JSON-ключ, сохранить рядом со скриптом как credentials.json
    3. Открыть таблицу → Поделиться → вставить email сервисного аккаунта (с правом редактора)
"""
 
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime
 
# ─── НАСТРОЙКИ ────────────────────────────────────────────────────────────────
 
SPREADSHEET_ID  = "10SLrg8hgNbWpBgrsuyozU0TLaArvO07UyPzo3T3L1x4"
SOURCE_SHEET    = "2026"
TARGET_SHEET    = "BD"
CREDENTIALS_FILE = "credentials.json"   # путь к JSON-ключу сервисного аккаунта
 
# Колонки пользователей в листе "2026" (0-based, col A = 0)
# Игорь=3, Марк=4, Аня=5, … Олег=25
USER_COLS = list(range(3, 26))
 
# Маппинг русских дней в английские
DAY_MAP = {
    "понедельник": "Monday",
    "вторник":     "Tuesday",
    "среда":       "Wednesday",
    "четверг":     "Thursday",
    "пятница":     "Friday",
    "суббота":     "Saturday",
    "воскресенье": "Sunday",
}
 
# ─── АВТОРИЗАЦИЯ ──────────────────────────────────────────────────────────────
 
def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    return gspread.authorize(creds)
 
 
# ─── ПАРСИНГ ЛИСТА "2026" ─────────────────────────────────────────────────────
 
def find_current_week_start(all_rows):
    """Возвращает индекс строки 'Забронировать' помеченной как текущая неделя."""
    for i, row in enumerate(all_rows):
        # Маркер текущей недели стоит в последней колонке строки Забронировать
        last = str(row[-1]).strip().lower() if row else ""
        if "текущ" in last:
            return i
    raise ValueError("Маркер текущей недели ('текущая неделя') не найден в листе 2026")
 
 
def find_week_label(all_rows, week_bron_idx):
    """Ищет строку-заголовок недели (например '1 неделя мая') выше строки Забронировать."""
    for i in range(week_bron_idx - 1, max(week_bron_idx - 5, -1), -1):
        row = all_rows[i]
        cell = str(row[1]).strip() if len(row) > 1 else ""
        if re.search(r'\d+\s+неделя', cell, re.IGNORECASE):
            return cell
    return ""
 
 
def get_user_names(all_rows):
    """Возвращает словарь {col_idx: clean_name} из строки-заголовка (строка 11, 0-based)."""
    header = all_rows[11] if len(all_rows) > 11 else []
    users = {}
    for col in USER_COLS:
        if col < len(header):
            raw = str(header[col]).strip()
            if raw and raw.lower() != "nan":
                # Убираем суффиксы уровня (LEADER, JUNIOR и т.д.)
                clean = re.split(r'\s+(LEADER|JUNIOR|NOT BAD|★|⭑|☆|Группа)', raw)[0].strip()
                users[col] = clean
    return users
 
 
def parse_week(all_rows, week_bron_idx, user_names):
    """
    Структура дня (5 строк, начиная со строки Забронировать):
      row+0: Забронировать  — bool по колонкам
      row+1: день / 'План'  — план по колонкам
      row+2: время/'нет тренировки' / 'Объем/м' — объём по колонкам
      row+3: дата / 'Коментарий' — комментарии по колонкам
      row+4: Состояние
    Затем следующий день снова начинается со строки Забронировать.
    """
    records = []
    row_idx = week_bron_idx   # первая строка = Забронировать для пн
 
    for _ in range(7):   # 7 дней
        if row_idx + 3 >= len(all_rows):
            break
 
        bron_row  = all_rows[row_idx]       # Забронировать
        plan_row  = all_rows[row_idx + 1]   # день / План
        vol_row   = all_rows[row_idx + 2]   # время / Объём
        com_row   = all_rows[row_idx + 3]   # дата / Комментарий
 
        # Определяем день недели из col[1] строки план
        day_ru  = str(plan_row[1]).strip().lower() if len(plan_row) > 1 else ""
        day_en  = DAY_MAP.get(day_ru, day_ru.capitalize())
 
        # Дата из col[1] строки комментариев
        date_str = str(com_row[1]).strip() if len(com_row) > 1 else ""
 
        # Время из col[1] строки объёма (общее для всех в этот день)
        time_raw = str(vol_row[1]).strip() if len(vol_row) > 1 else ""
        is_no_train = "нет тренировки" in time_raw.lower()
        is_remote_day = "удал" in time_raw.lower()
 
        # Нормализуем время: "20 00" → "20:00"
        time_clean = ""
        if not is_no_train and not is_remote_day and time_raw:
            time_clean = re.sub(r'(\d{1,2})\s+(\d{2})', r'\1:\2', time_raw)
 
        for col, user in user_names.items():
            if col >= len(bron_row):
                continue
 
            booked  = str(bron_row[col]).strip().upper() == "TRUE"
            plan    = str(plan_row[col]).strip()  if col < len(plan_row)  else ""
            volume  = str(vol_row[col]).strip()   if col < len(vol_row)   else ""
            comment = str(com_row[col]).strip()   if col < len(com_row)   else ""
 
            # Чистим "nan"
            plan    = "" if plan    in ("nan", "None") else plan
            volume  = "" if volume  in ("nan", "None") else volume
            comment = "" if comment in ("nan", "None") else comment
 
            remote = is_remote_day or "удал" in plan.lower() or "удал" in comment.lower()
 
            records.append({
                "User":     user,
                "Date":     date_str,
                "Day":      day_en,
                "Time":     time_clean,
                "Plan":     plan,
                "Volume":   volume,
                "Remote":   "Yes" if remote else "No",
                "Booked":   "TRUE" if booked else "FALSE",
                "Comments": comment,
            })
 
        # Следующий день: пропускаем Состояние (row+4) и идём к следующему Забронировать
        row_idx += 5
 
    return records
 
 
# ─── ЗАПИСЬ В ЛИСТ "bd" ───────────────────────────────────────────────────────
 
HEADER = ["User", "Date", "Day", "Time", "Plan", "Volume", "Remote", "Booked", "Comments"]
 
def write_to_bd(client, records, week_label, week_dates):
    """Очищает лист bd и записывает заголовок + данные."""
    ss = client.open_by_key(SPREADSHEET_ID)
 
    # Ищем лист bd (регистронезависимо, обрезаем пробелы)
    bd = None
    for ws in ss.worksheets():
        if ws.title.strip().lower() == TARGET_SHEET.strip().lower():
            bd = ws
            break
    if bd is None:
        bd = ss.add_worksheet(title=TARGET_SHEET, rows=2000, cols=20)
 
    bd.clear()
 
    # Строка 1: мета-инфо о неделе
    bd.update("A1", [[f"{week_label}  |  {week_dates}"]])
 
    # Строка 2: заголовки колонок
    bd.update("A2", [HEADER])
 
    # Строки 3+: данные
    rows = [[r[k] for k in HEADER] for r in records]
    if rows:
        bd.update(f"A3", rows)
 
    print(f"✓ Записано {len(rows)} строк в лист '{TARGET_SHEET}'")
    print(f"  Неделя: {week_label}  |  {week_dates}")
 
 
# ─── MAIN ─────────────────────────────────────────────────────────────────────
 
def main():
    print("Подключаемся к Google Sheets...")
    client = get_client()
 
    ss = client.open_by_key(SPREADSHEET_ID)
    src = ss.worksheet(SOURCE_SHEET)
 
    print(f"Читаем лист '{SOURCE_SHEET}'...")
    all_rows = src.get_all_values()   # список списков строк
 
    # Находим текущую неделю
    week_bron_idx = find_current_week_start(all_rows)
    print(f"  Найдена текущая неделя: строка {week_bron_idx + 1}")
 
    week_label = find_week_label(all_rows, week_bron_idx)
    print(f"  Заголовок недели: «{week_label}»")
 
    user_names = get_user_names(all_rows)
    print(f"  Учеников: {len(user_names)}")
 
    # Парсим 7 дней
    records = parse_week(all_rows, week_bron_idx, user_names)
 
    # Диапазон дат для заголовка
    dates = [r["Date"] for r in records if r["Date"]]
    week_dates = f"{dates[0]} — {dates[-1]}" if len(dates) >= 2 else ""
 
    # Пишем в bd
    all_sheets = [ws.title for ws in ss.worksheets()]
    print(f"\nЛисты в таблице: {all_sheets}")
    print(f"Записываем в лист '{TARGET_SHEET}'...")
    write_to_bd(client, records, week_label, week_dates)
    print("\nГотово!")
 
 
if __name__ == "__main__":
    main()