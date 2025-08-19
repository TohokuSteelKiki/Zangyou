from tkinter import messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoAlertPresentException, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import datetime
import time
import sys
import tkinter as tk
from tkinter import simpledialog
import os  # 追加


# ====== ユーザー入力（パスワード・理由） ======
root = tk.Tk()
root.withdraw()

# --- パスワード入力（常に必要） ---
PASSWORD = simpledialog.askstring(
    "パスワード入力", "ログイン用パスワードを入力してください：", show="*"
)
if not PASSWORD:
    print("[ERROR] パスワードが入力されませんでした。")
    sys.exit(1)

# --- 残業申請を実行するか確認 ---
proceed = messagebox.askyesno("確認", "残業申請を実行しますか？")
if not proceed:
    print("[INFO] ユーザーが申請をキャンセルしました。")
# sys.exit(0)
if proceed:
    # --- 残業理由の入力（申請する場合のみ） ---
    ZANGYO_REASON = simpledialog.askstring(
        "残業理由入力", "残業申請の理由を入力してください："
    )
    if not ZANGYO_REASON:
        print("[ERROR] 残業理由が入力されませんでした。")
        sys.exit(1)


# ====== 定時設定 ======
定時 = datetime.datetime.strptime("17:00", "%H:%M")
# ====== 設定 ======
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))  # スクリプトの絶対パスを取得
EXCEL_PATH = os.path.join(
    SCRIPT_DIR, "IDPASS.xlsx"
)  # 同じフォルダ内のExcelファイルを指定
TARGET_SCRIPT = "TimeProGX"
LOGIN_URL = "http://128.198.11.125/xgweb/login.asp"

# ====== ログインID取得（Excelから） ======
try:
    df = pd.read_excel(EXCEL_PATH, dtype={"ID": str})  # ← dtype指定で文字列として読む
    row = df[df["スクリプト"] == TARGET_SCRIPT].iloc[0]
    LOGIN_ID = row["ID"].strip()  # strip()で空白除去も安全に
except Exception as e:
    print(f"[ERROR] Excel読み込み失敗: {e}")
    sys.exit(1)

# ====== Chrome起動オプション ======
options = Options()
# options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=options)
driver.implicitly_wait(3)

try:
    print("[INFO] ログインページにアクセス中...")
    driver.get(LOGIN_URL)
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, "LoginID")))

    driver.find_element(By.NAME, "LoginID").send_keys(LOGIN_ID)
    driver.find_element(By.NAME, "PassWord").send_keys(PASSWORD)
    driver.find_element(By.NAME, "btnLogin").click()

    WebDriverWait(driver, 5).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "frame"))
    )

    # ====== 退勤ボタン探索 ======
    frames = driver.find_elements(By.TAG_NAME, "frame")
    found = False
    for i in range(len(frames)):
        driver.switch_to.default_content()
        driver.switch_to.frame(i)
        try:
            retire_button = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "退　勤"))
            )
            retire_button.click()
            print(f"[SUCCESS] 退勤ボタンを Frame {i} 内でクリックしました。")
            found = True
            break
        except Exception:
            continue

    if not found:
        print("[WARNING] 退勤リンクが見つかりませんでした。")
        driver.quit()
        sys.exit(1)

    # ====== ポップアップ切替 ======
    main_window = driver.current_window_handle
    WebDriverWait(driver, 5).until(lambda d: len(d.window_handles) > 1)
    for handle in driver.window_handles:
        if handle != main_window:
            driver.switch_to.window(handle)
            break

    # ====== 打刻時間取得 ======
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ap_table"))
    )
    rows = driver.find_elements(By.CSS_SELECTOR, ".ap_table tr")
    punch_time = None
    for row in rows:
        tds = row.find_elements(By.TAG_NAME, "td")
        if len(tds) == 2 and "打刻" in tds[0].text:
            punch_time = tds[1].text.strip()
            break

    if not punch_time:
        print("[WARNING] 打刻時間が取得できませんでした。")
        driver.quit()
        sys.exit(1)

    print(f"[INFO] 打刻時間: {punch_time}")

    # ====== ポップアップ閉じる（戻る） ======
    try:
        driver.find_element(By.LINK_TEXT, "戻る").click()
        print("[INFO] ポップを閉じました。")
        WebDriverWait(driver, 3).until(lambda d: len(d.window_handles) == 1)
    except Exception as e:
        print(f"[WARNING] 戻るボタン操作失敗: {e}")

    driver.switch_to.window(main_window)
    driver.switch_to.default_content()

    if not proceed:
        print("[INFO] 残業申請しないので終了します。")
        sys.exit(0)

    # ====== 残業時間判定 ======
    punch_dt = datetime.datetime.strptime(punch_time, "%H:%M")
    delta_min = (punch_dt - 定時).total_seconds() / 60
    # if delta_min < 10:
    #     print("[INFO] 残業時間が10分未満のため申請をスキップします。")
    #     driver.quit()
    #     sys.exit(0)

    start_time = 定時.strftime("%H:%M")
    end_time = punch_time
    print(f"[INFO] 残業申請時間: {start_time} ～ {end_time}")

    # ====== メニュー遷移（frameTop → frameBtm） ======
    WebDriverWait(driver, 5).until(
        EC.frame_to_be_available_and_switch_to_it("frameTop")
    )
    WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "届出処理"))
    ).click()

    driver.switch_to.default_content()
    WebDriverWait(driver, 5).until(
        EC.frame_to_be_available_and_switch_to_it("frameBtm")
    )
    WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//span[contains(text(), '就業届出処理')]")
        )
    ).click()
    WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//span[.//img[contains(@alt, '時間外申請')]]")
        )
    ).click()

    # ====== 申請画面で入力 ======
    driver.switch_to.default_content()
    WebDriverWait(driver, 5).until(
        EC.frame_to_be_available_and_switch_to_it("frameBtm")
    )

    driver.find_element(By.ID, "TxtExtStart0").clear()
    driver.find_element(By.ID, "TxtExtStart0").send_keys(start_time)
    driver.find_element(By.ID, "TxtExtEnd0").clear()
    driver.find_element(By.ID, "TxtExtEnd0").send_keys(end_time)
    driver.find_element(By.ID, "TxtNotes0").clear()
    driver.find_element(By.ID, "TxtNotes0").send_keys(ZANGYO_REASON)

    checkbox = driver.find_element(By.ID, "ChkExtNotrpt0")
    if checkbox.is_selected():
        checkbox.click()

    print("[SUCCESS] 残業申請フォーム入力完了")

    # --- 登録ボタンを押す ---
    apply_button = driver.find_element(
        By.XPATH, "//input[@name='ActBtn' and @value='登録']"
    )
    apply_button.click()

    # --- 登録ポップアップに自動応答 ---
    try:
        WebDriverWait(driver, 10).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        print(f"✅ 登録ポップアップ: {alert.text}")
        alert.accept()
        print("→ OKボタンをクリックしました。")
    except TimeoutException:
        print("⚠️ アラートが表示されませんでした。")

    # 申請内容確認
    try:
        # frameTopで「届出処理」を再クリック（アクティブ化）
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it("frameTop")
        )
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "届出処理"))
        ).click()

        # frameBtmで「届出データ表示」をクリック
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it("frameBtm")
        )
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//span[.//img[contains(@alt, '届出データ表示')]]")
            )
        ).click()

        time.sleep(2)  # 表示待ち（必要なら明示）

        # 表示ページのframeBtmに再度切り替えて最下部へスクロール
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it("frameBtm")
        )
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        print("[INFO] 届出データ表示ページを表示し、スクロール完了。")

    except Exception as e:
        print(f"[WARNING] 届出データ表示の確認中にエラーが発生しました: {e}")

except Exception as e:
    print(f"[ERROR] 処理中にエラーが発生しました: {e}")

finally:
    # driver.quit()
    #print("[INFO] ブラウザを閉じて終了しました。")
    print("[INFO] 終了")
