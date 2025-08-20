from tkinter import messagebox
from tkinter import simpledialog
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


from selenium.webdriver.support import expected_conditions as EC


import tkinter as tk

import os  # 追加

# ====== GUI入力（パスワード・残業理由） ======
root = tk.Tk()
root.withdraw()
LOGIN_ID = "youID"
PASSWORD = simpledialog.askstring(
    "パスワード入力", "ログイン用パスワードを入力してください：", show="*"
)
if not PASSWORD:
    print("[ERROR] パスワードが入力されませんでした。")
    sys.exit(1)



TARGET_SCRIPT = "TimeProGX"
LOGIN_URL = "http://128.198.11.125/xgweb/login.asp"


# ====== Chrome起動 ======
options = Options()
# options.add_argument("--headless")  # GUI確認したければコメントアウト
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
driver = webdriver.Chrome(service=Service(), options=options)
driver.implicitly_wait(3)


print("[INFO] ログインページにアクセス中...")
driver.get(LOGIN_URL)

# ログインフォームの表示を待機
WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, "LoginID")))

# 値の送信
driver.find_element(By.NAME, "LoginID").send_keys(LOGIN_ID)
driver.find_element(By.NAME, "PassWord").send_keys(PASSWORD)
print("[INFO] ログイン情報を入力、ログインボタンクリック")
driver.find_element(By.NAME, "btnLogin").click()

# 少し待機
time.sleep(2)

# フレームがあるかどうか確認
frames = driver.find_elements(By.TAG_NAME, "frame")
print(f"[INFO] ログイン後のフレーム数: {len(frames)}")

if len(frames) == 0:
    with open("login_debug.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    driver.save_screenshot("login_error.png")
    print(
        "[WARN] フレームが検出されませんでした。HTMLとスクリーンショットを保存しました。"
    )
    raise Exception("ログイン失敗または画面構造の変更")

print("[SUCCESS] ログイン成功！次の処理へ進めます")


from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ====== frameTop: 「就業情報」クリック ======
driver.switch_to.default_content()
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("frameTop"))
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "就業情報"))).click()

# ====== frameBtm: 「就業日次処理」→「就業週報」クリック ======
driver.switch_to.default_content()
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("frameBtm"))

# 「就業日次処理」クリック（<span>タグ内のテキスト）
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '就業日次処理')]"))
).click()

# 「就業週報」クリック（画像+テキスト要素）
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//span[.//img[contains(@alt, '就業週報')]]"))
).click()

# frameBtm に切り替えた後にスクロール
driver.switch_to.default_content()
WebDriverWait(driver, 10).until(
    EC.frame_to_be_available_and_switch_to_it("frameBtm")
)
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(1)

# 所定日数・出勤日数・年休日数を取得
data_map = {"所定日数": None, "出勤日数": None, "年休日数": None}

cells = driver.find_elements(By.XPATH, "//tr[contains(@class, 'ap_tr_base')]/td")
for i in range(len(cells) - 1):
    label = cells[i].text.strip()
    if label in data_map:
        data_map[label] = cells[i + 1].text.strip()

# 結果出力
print(f"📊 所定日数: {data_map['所定日数']}, 出勤日数: {data_map['出勤日数']}, 年休日数: {data_map['年休日数']}")
