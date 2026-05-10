import gspread
from google.oauth2.service_account import Credentials

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
client = gspread.authorize(creds)

SPREADSHEET_ID = "10SLrg8hgNbWpBgrsuyozU0TLaArvO07UyPzo3T3L1x4"

try:
    sh = client.open_by_key(SPREADSHEET_ID)
    print("✅ Подключение успешно! Название таблицы:", sh.title)
except Exception as e:
    print("❌ Ошибка подключения:", e)