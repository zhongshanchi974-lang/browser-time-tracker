import time
import psutil
import win32gui
import win32process
from collections import defaultdict
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib as mpl
from datetime import datetime, timedelta
import csv
import os
import re

font_path = "C:/Windows/Fonts/msgothic.ttc"
font_prop = fm.FontProperties(fname=font_path)
mpl.rcParams["font.family"] = font_prop.get_name()

INTERVAL = 5  # 秒
BROWSERS = ["chrome.exe", "msedge.exe", "firefox.exe"]

daily_data = defaultdict(lambda: defaultdict(int))

# 既知サイトの判定ルール（順番に評価）
KNOWN_SITES = [
    ("youtube",       lambda t: "YouTube" in t),
    ("google",        lambda t: "Google" in t),
    ("github",        lambda t: "GitHub" in t),
    ("qiita",         lambda t: "Qiita" in t),
    ("stackoverflow", lambda t: "Stack Overflow" in t),
]

def get_site_from_title(title: str) -> str:
    """タイトルからサイト名を返す。既知サイトはラベル名、未知サイトはタイトルから抽出。"""

    # 既知サイト判定
    for label, match in KNOWN_SITES:
        if match(title):
            return label

    # ブラウザのタイトルは多くの場合 "ページタイトル - サイト名" や "ページタイトル | サイト名" の形式
    # 末尾のセパレータ以降をサイト名として使う
    sep_match = re.search(r'[\-\|–—]\s*([^\-\|–—]+?)\s*$', title)
    if sep_match:
        site_name = sep_match.group(1).strip()
        if site_name:
            return site_name

    # セパレータがない場合はタイトル全体を短く切って使う
    return title[:30] if title else "unknown"


def get_active_window():
    hwnd = win32gui.GetForegroundWindow()
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['pid'] == pid:
            return proc.info['name'], win32gui.GetWindowText(hwnd)
    return None, None


print("1日ごとのアクセス時間計測を開始（Ctrl+Cで終了）")

try:
    while True:
        today = datetime.now().strftime("%Y-%m-%d")
        proc_name, title = get_active_window()

        if proc_name and proc_name.lower() in BROWSERS:
            site = get_site_from_title(title)  # ← ここを変更
            daily_data[today][site] += INTERVAL

        time.sleep(INTERVAL)

except KeyboardInterrupt:
    print("\n計測終了")

    # ===== CSV保存 =====
    os.makedirs("data", exist_ok=True)

    for date, sites in daily_data.items():
        with open(f"data/{date}.csv", "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["site", "seconds"])
            for site, sec in sites.items():
                writer.writerow([site, sec])

    # ===== 今日の円グラフ =====
    today = datetime.now().strftime("%Y-%m-%d")
    sites = daily_data[today]
    total = sum(sites.values())

    if total > 0:
        plt.figure(figsize=(8, 8))
        plt.pie(
            [sec / total * 100 for sec in sites.values()],
            labels=sites.keys(),
            autopct="%1.1f%%",
            startangle=140
        )
        plt.title(f"{today} サイト別アクセス割合")
        plt.show()
    else:
        print("今日のブラウザ使用がありません")

    # ===== 直近1週間の棒グラフ =====
    weekly_site_time = defaultdict(int)
    today_date = datetime.now().date()

    for i in range(7):
        day = today_date - timedelta(days=i)
        path = f"data/{day}.csv"

        if os.path.exists(path):
            with open(path, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    weekly_site_time[row["site"]] += int(row["seconds"])

    if weekly_site_time:
        plt.figure(figsize=(10, 6))
        plt.bar(
            weekly_site_time.keys(),
            [sec / 3600 for sec in weekly_site_time.values()]
        )
        plt.xticks(rotation=45, ha="right")
        plt.ylabel("利用時間（時間）")
        plt.title("直近1週間 サイト別アクセス時間")
        plt.tight_layout()
        plt.show()
    else:
        print("1週間分のデータがありません")