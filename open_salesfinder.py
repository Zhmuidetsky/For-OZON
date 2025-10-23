import subprocess
import shutil
import platform
from datetime import date, timedelta
from pathlib import Path

# === НОМЕРА КАТЕГОРИЙ ===
category_numbers = [
    1, 4, 7, 10, 13, 16, 19, 7711,
    2, 5, 8, 11, 14, 17, 20, 7788,
    3, 6, 9, 12, 15, 18, 21, 7789,
]

# === ДАТЫ ПРОШЛОЙ НЕДЕЛИ ===
today = date.today()
start_of_week = today - timedelta(days=today.weekday())
last_week_monday = start_of_week - timedelta(days=7)
last_week_sunday = start_of_week - timedelta(days=1)

d1 = last_week_monday.strftime('%Y-%m-%d')
d2 = last_week_sunday.strftime('%Y-%m-%d')

# === ССЫЛКИ ===
base = "https://salesfinder.ru/ozon/category/{cat}/info/products?date={d1}&date2={d2}"
urls = [base.format(cat=c, d1=d1, d2=d2) for c in category_numbers]

# === ПОИСК CHROME ===
def find_chrome() -> str | None:
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for p in candidates:
        if Path(p).exists():
            return p
    return shutil.which("chrome") or shutil.which("google-chrome")

chrome_path = find_chrome()

if chrome_path:
    subprocess.Popen([chrome_path, "--new-window"] + urls)
else:
    raise FileNotFoundError("Google Chrome не найден. Пропишите путь вручную в chrome_path.")
