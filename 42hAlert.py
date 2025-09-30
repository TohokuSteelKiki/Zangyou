import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
import os
import sys
import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException, TimeoutException

# ====== 残業判定の定数 ======
ZANGYOU_LIMIT_HOUR = 2
ZANGYOU_LIMIT_MINUTES = ZANGYOU_LIMIT_HOUR * 60  # 40時間 = 2400分
ZANGYOU_ALERT_DAY = 2  # 月の20日以降で警告判定
# ====== GUI入力（パスワード） ======
root = tk.Tk()
root.withdraw()
LOGIN_ID = "3046"
PASSWORD = simpledialog.askstring(
    "パスワード入力", "ログイン用パスワードを入力してください：", show="*"
)
if not PASSWORD:
    print("[ERROR] パスワードが入力されませんでした。")
    sys.exit(1)

# ====== ログインURLなど ======
LOGIN_URL = "http://128.198.11.125/xgweb/login.asp"

# ====== Chrome起動 ======
options = Options()
# options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
driver = webdriver.Chrome(service=Service(), options=options)
driver.implicitly_wait(3)

print("[INFO] ログインページにアクセス中...")
driver.get(LOGIN_URL)

WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, "LoginID")))
driver.find_element(By.NAME, "LoginID").send_keys(LOGIN_ID)
driver.find_element(By.NAME, "PassWord").send_keys(PASSWORD)
driver.find_element(By.NAME, "btnLogin").click()
time.sleep(2)

frames = driver.find_elements(By.TAG_NAME, "frame")
if len(frames) == 0:
    print("[ERROR] フレームが検出されません。ログイン失敗の可能性")
    sys.exit(1)

print("[SUCCESS] ログイン成功！")

# ====== 就業週報ページへ遷移 ======
driver.switch_to.default_content()
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("frameTop"))
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.LINK_TEXT, "就業情報"))
).click()

driver.switch_to.default_content()
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("frameBtm"))
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '就業日次処理')]"))
).click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//span[.//img[contains(@alt, '就業週報')]]"))
).click()

# スクロールしてデータを描画させる
driver.switch_to.default_content()
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("frameBtm"))
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(1)

# ====== 所定日数・出勤日数・年休日数を取得 ======
data_map = {"所定日数": None, "出勤日数": None, "年休日数": None}
cells = driver.find_elements(By.XPATH, "//tr[contains(@class, 'ap_tr_base')]/td")
for i in range(len(cells) - 1):
    label = cells[i].text.strip()
    if label in data_map:
        data_map[label] = cells[i + 1].text.strip()

print(
    f"📊 所定日数: {data_map['所定日数']}, 出勤日数: {data_map['出勤日数']}, 年休日数: {data_map['年休日数']}"
)

# ====== 早出残業（合計）取得 ======
overtime_total_elem = driver.find_element(By.XPATH, "//td[@title='合計    早出残業']")
early_overtime_total = overtime_total_elem.text.strip()
print(f"🕒 早出残業合計: {early_overtime_total}")


# ====== 計算処理 ======
def time_str_to_minutes(timestr):
    """ "H:MM" → 分に変換"""
    try:
        hours, minutes = map(int, timestr.split(":"))
        return hours * 60 + minutes
    except:
        return 0


def minutes_to_time_str(minutes):
    """分 → "H:MM" 文字列"""
    h = minutes // 60
    m = minutes % 60
    return f"{h}:{m:02d}"


try:
    early_total_min = time_str_to_minutes(early_overtime_total)
    work_days = float(data_map["出勤日数"])
    planned_days = float(data_map["所定日数"])
    remaining_days = planned_days - work_days

    # 平均と予測
    avg_overtime_min = early_total_min / work_days if work_days else 0
    projected_total_min = avg_overtime_min * planned_days

    print(f"\n【📈 残業予測】")
    print(f"- 平均残業時間/日: {minutes_to_time_str(int(avg_overtime_min))}")
    print(f"- 残業時間予測（月末）: {minutes_to_time_str(int(projected_total_min))}")
    print(f"- 月の残り出勤数: {remaining_days:.1f} 日")

    # ====== 残業警告ロジック ======
    today = datetime.datetime.today()
    show_alert = False

    if today.day >= ZANGYOU_ALERT_DAY and projected_total_min >= ZANGYOU_LIMIT_MINUTES:
        show_alert = True

    if show_alert:
        messagebox.showwarning(
            "⚠️ 残業時間注意",
            f"このままでは月末の残業時間が{ZANGYOU_LIMIT_MINUTES // 60}時間を超えます！\n予測: {minutes_to_time_str(int(projected_total_min))}",
        )
    else:
        print(
            "\n✔️ 残業時間{ZANGYOU_LIMIT_HOUR}h超の可能性は低い、{ZANGYOU_ALERT_DAY}日未満のため通知はしません。"
        )

except Exception as e:
    print(f"[ERROR] 残業時間の計算中にエラー: {e}")
