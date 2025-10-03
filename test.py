# -*- coding: utf-8 -*-
"""
TimeProGX 残業申請 + 残業予測
- UI入力: Tkinter
- 自動操作: Selenium Edge
- ログインID: Excel参照
- PyInstaller配布を考慮した WebDriver 解決
"""

from __future__ import annotations

# ====== 標準 ======
import datetime as dt
import os
import sys
import time
from pathlib import Path
from typing import Optional, Dict, Tuple

# ====== GUI ======
import tkinter as tk
from tkinter import messagebox, font

# ====== 表 ======
import pandas as pd

# ====== Selenium ======
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    UnexpectedAlertPresentException,
)
from selenium.webdriver.common.alert import Alert

# =============================================================================
# 定数
# =============================================================================

APP_NAME = "TimeProGX"
LOGIN_URL = "http://128.198.11.125/xgweb/login.asp"
EXCEL_FILENAME = "ID.xlsx"
EXCEL_COL_KEY = "項目"
EXCEL_COL_VAL = "定数"
EXCEL_LOGIN_KEYS = [
    "ID",
    "ログインID",
    "LoginID",
    "login_id",
    "TimeProGX（社員コード）",
]
EXCEL_FIXED_OFF_KEYS = ["定時", "退社基準", "FIXED_OFF_TIME", "終業時刻"]


# テストモード: True=登録クリックしない / False=登録クリックする（本番）
IS_TEST: bool = True

# 定時
FIXED_OFF_TIME = dt.datetime.strptime("17:00", "%H:%M")  # 退社基準

# 残業アラート
ZANGYOU_LIMIT_HOUR = 30  # 月末予測がこの時間以上で警告
ZANGYOU_ALERT_DAY = 20  # 月内のこの日以降に判定

# 入力仕様
MAX_REASON_LEN = 20

# ブラウザ自動終了（秒）
BROWSER_AUTO_CLOSE_AFTER_SEC = 300  # 5分

# =============================================================================
# 汎用ユーティリティ
# =============================================================================


def log(info: str) -> None:
    print(f"[INFO] {info}")


def warn(msg: str) -> None:
    print(f"[WARNING] {msg}")


def err(msg: str) -> None:
    print(f"[ERROR] {msg}")


def time_str_to_minutes(timestr: str) -> int:
    try:
        h, m = map(int, timestr.strip().split(":"))
        return h * 60 + m
    except Exception:
        return 0


def minutes_to_time_str(minutes: int) -> str:
    h = max(minutes, 0) // 60
    m = max(minutes, 0) % 60
    return f"{h}:{m:02d}"


# =============================================================================
# Tk ダイアログ
# =============================================================================


def custom_input_dialog(
    title: str, prompt: str, show: Optional[str] = None, maxlen: Optional[int] = None
) -> Optional[str]:
    win = tk.Toplevel()
    win.title(title)
    win.geometry("420x160")
    win.resizable(False, False)
    win.grab_set()

    label_font = font.Font(size=12)
    entry_font = font.Font(size=12)

    tk.Label(
        win, text=prompt, font=label_font, anchor="w", justify="left", wraplength=380
    ).pack(padx=14, pady=(12, 6), fill="x")
    entry = tk.Entry(win, font=entry_font, width=30, show=show)
    entry.pack(padx=14, fill="x")
    entry.focus_set()

    result = {"value": None}

    def submit():
        text = entry.get()
        if maxlen is not None and len(text) > maxlen:
            messagebox.showwarning(
                "文字数制限", f"{maxlen}文字以内で入力してください。"
            )
            return
        result["value"] = text
        win.destroy()

    def cancel():
        win.destroy()

    btn = tk.Frame(win)
    btn.pack(pady=12)
    tk.Button(btn, text="OK", font=entry_font, width=10, command=submit).pack(
        side=tk.LEFT, padx=8
    )
    tk.Button(btn, text="キャンセル", font=entry_font, width=10, command=cancel).pack(
        side=tk.LEFT, padx=8
    )

    win.bind("<Return>", lambda _: submit())
    win.bind("<Escape>", lambda _: cancel())
    win.wait_window()
    return result["value"]


def ask_password_and_reason() -> Tuple[str, Optional[str]]:
    root = tk.Tk()
    root.withdraw()

    password = custom_input_dialog(
        "パスワード入力", "ログイン用パスワードを入力してください：", show="*"
    )
    if not password:
        err("パスワードが入力されませんでした。")
        sys.exit(1)

    proceed = messagebox.askyesno("確認", "残業申請を実行しますか？")
    reason = None
    if proceed:
        while True:
            reason = custom_input_dialog(
                "残業理由入力",
                f"残業申請の理由を入力してください（{MAX_REASON_LEN}文字以内）：",
                maxlen=MAX_REASON_LEN,
            )
            if reason is None:
                err("残業理由が入力されませんでした。")
                sys.exit(1)
            if len(reason.strip()) == 0:
                messagebox.showwarning(
                    "入力エラー", "理由が空白です。内容を入力してください。"
                )
                continue
            break

    return password, reason


# =============================================================================
# データ取得
# =============================================================================


def _load_excel_kv(excel_path: Path) -> Dict[str, str]:
    try:
        df = pd.read_excel(excel_path, dtype=str)
        if EXCEL_COL_KEY not in df.columns or EXCEL_COL_VAL not in df.columns:
            raise RuntimeError(
                f"Excelに必要な列がありません: {EXCEL_COL_KEY}, {EXCEL_COL_VAL}"
            )
        kv = {}
        for _, row in df.iterrows():
            k = str(row[EXCEL_COL_KEY]).strip()
            v = "" if pd.isna(row[EXCEL_COL_VAL]) else str(row[EXCEL_COL_VAL]).strip()
            if k:
                kv[k] = v
        return kv
    except Exception as e:
        raise RuntimeError(f"Excel読み込み失敗: {e}")


def _get_from_kv(
    kv: Dict[str, str], candidates: list[str], *, required: bool = False
) -> Optional[str]:
    for c in candidates:
        if c in kv and str(kv[c]).strip():
            return str(kv[c]).strip()
    if required:
        raise RuntimeError(f"Excelに必要キーが見つかりません: {candidates}")
    return None


def parse_hhmm(s: str) -> str:
    t = str(s).strip().replace("：", ":")
    if not t:
        raise ValueError("empty time")
    parts = t.split(":")
    if len(parts) == 3:  # HH:MM:SS → HH:MM
        h, m, _ = parts
    elif len(parts) == 2:  # HH:MM
        h, m = parts
    else:
        # "1700" 等にも一応対応
        digits = "".join(ch for ch in t if ch.isdigit())
        if len(digits) == 4:
            h, m = digits[:2], digits[2:]
        else:
            raise ValueError(f"unsupported time format: {s}")
    h = f"{int(h):02d}"
    m = f"{int(m):02d}"
    return f"{h}:{m}"


def resolve_driver_path() -> Path:
    # 同階層に msedgedriver.exe を配置する前提（PyInstaller対応）
    driver_filename = "msedgedriver.exe"
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).parent
    elif hasattr(sys, "_MEIPASS"):
        base_dir = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    else:
        base_dir = Path(__file__).resolve().parent
    return base_dir / driver_filename


# =============================================================================
# Selenium 操作
# =============================================================================


def handle_possible_alert(drv: webdriver.Edge, timeout: int = 0) -> bool:
    """アラートがあれば受理して True を返す"""
    try:
        WebDriverWait(drv, timeout).until(EC.alert_is_present())
        a = drv.switch_to.alert
        text = a.text
        a.accept()
        log(f"アラート自動処理: {text}")
        return True
    except TimeoutException:
        return False
    except Exception as e:
        warn(f"アラート処理失敗: {e}")
        return False


def create_driver(driver_path: Path) -> webdriver.Edge:
    if not driver_path.exists():
        messagebox.showerror(
            "ドライバー未検出",
            f"WebDriver が見つかりません:\n{driver_path}\nEXE と同じフォルダに msedgedriver.exe を配置してください。",
        )
        sys.exit(1)

    options = EdgeOptions()
    # options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    # 未処理アラートは自動 accept
    options.set_capability("unhandledPromptBehavior", "accept")

    service = EdgeService(executable_path=str(driver_path))
    drv = webdriver.Edge(service=service, options=options)
    drv.implicitly_wait(3)
    return drv


def wait(drv: webdriver.Edge, timeout: int = 10) -> WebDriverWait:
    return WebDriverWait(drv, timeout)


def find_and_click_in_frames(
    drv: webdriver.Edge, by: By, value: str, click: bool = True, frame_wait: int = 3
) -> bool:
    """複数frame/iframeを総当たりして最初に見つかった要素をクリックする"""
    # 事前にアラート掃除
    handle_possible_alert(drv, timeout=1)

    frames = drv.find_elements(By.TAG_NAME, "frame")
    if not frames:
        frames = drv.find_elements(By.TAG_NAME, "iframe")

    for i in range(len(frames)):
        try:
            drv.switch_to.default_content()
            drv.switch_to.frame(i)
            elem = WebDriverWait(drv, frame_wait).until(
                EC.element_to_be_clickable((by, value))
            )
            if click:
                elem.click()
            return True
        except UnexpectedAlertPresentException:
            handle_possible_alert(drv, timeout=2)
            continue
        except Exception:
            continue

    # デフォルトに戻す
    drv.switch_to.default_content()
    return False


def switch_to_new_window(drv: webdriver.Edge, timeout: int = 5) -> None:
    main = drv.current_window_handle
    WebDriverWait(drv, timeout).until(lambda d: len(d.window_handles) > 1)
    for h in drv.window_handles:
        if h != main:
            drv.switch_to.window(h)
            return


def get_punch_time_from_popup(drv: webdriver.Edge) -> Optional[str]:
    wait(drv, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "ap_table")))
    rows = drv.find_elements(By.CSS_SELECTOR, ".ap_table tr")
    for r in rows:
        tds = r.find_elements(By.TAG_NAME, "td")
        if len(tds) == 2 and "打刻" in tds[0].text:
            return tds[1].text.strip()
    return None


def navigate_menu_to_overtime_form(drv: webdriver.Edge) -> None:
    wait(drv, 5).until(EC.frame_to_be_available_and_switch_to_it("frameTop"))
    wait(drv, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, "届出処理"))).click()

    drv.switch_to.default_content()
    wait(drv, 5).until(EC.frame_to_be_available_and_switch_to_it("frameBtm"))
    wait(drv, 5).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//span[contains(text(), '就業届出処理')]")
        )
    ).click()
    wait(drv, 5).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//span[.//img[contains(@alt, '時間外申請')]]")
        )
    ).click()


def fill_overtime_form(
    drv: webdriver.Edge, start_hm: str, end_hm: str, reason: str
) -> None:
    drv.switch_to.default_content()
    wait(drv, 5).until(EC.frame_to_be_available_and_switch_to_it("frameBtm"))

    def set_value(elem_id: str, value: str) -> None:
        e = drv.find_element(By.ID, elem_id)
        e.clear()
        e.send_keys(value)

    set_value("TxtExtStart0", start_hm)
    set_value("TxtExtEnd0", end_hm)
    set_value("TxtNotes0", reason)

    checkbox = drv.find_element(By.ID, "ChkExtNotrpt0")
    if checkbox.is_selected():
        checkbox.click()

    log("残業申請フォーム入力完了")

    apply_button = drv.find_element(
        By.XPATH, "//input[@name='ActBtn' and @value='登録']"
    )
    if not IS_TEST:
        apply_button.click()

    try:
        wait(drv, 10).until(EC.alert_is_present())
        alert = drv.switch_to.alert
        print(f"✅ 登録ポップアップ: {alert.text}")
        alert.accept()
        print("→ OKボタンをクリックしました。")
    except TimeoutException:
        warn("アラートが表示されませんでした。")


def navigate_to_weekly_report(drv: webdriver.Edge) -> None:
    drv.switch_to.default_content()
    wait(drv, 10).until(EC.frame_to_be_available_and_switch_to_it("frameTop"))
    wait(drv, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "就業情報"))).click()

    drv.switch_to.default_content()
    wait(drv, 10).until(EC.frame_to_be_available_and_switch_to_it("frameBtm"))
    wait(drv, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//span[contains(text(), '就業日次処理')]")
        )
    ).click()
    wait(drv, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//span[.//img[contains(@alt, '就業週報')]]")
        )
    ).click()

    drv.switch_to.default_content()
    wait(drv, 10).until(EC.frame_to_be_available_and_switch_to_it("frameBtm"))
    drv.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)


def extract_weekly_metrics(drv: webdriver.Edge) -> Dict[str, str]:
    data_map = {"所定日数": None, "出勤日数": None, "年休日数": None}
    cells = drv.find_elements(By.XPATH, "//tr[contains(@class, 'ap_tr_base')]/td")
    for i in range(len(cells) - 1):
        label = cells[i].text.strip()
        if label in data_map:
            data_map[label] = cells[i + 1].text.strip()

    early_total_text = drv.find_element(
        By.XPATH, "//td[@title='合計    早出残業']"
    ).text.strip()

    return {
        "所定日数": data_map["所定日数"] or "0",
        "出勤日数": data_map["出勤日数"] or "0",
        "年休日数": data_map["年休日数"] or "0",
        "早出残業合計": early_total_text or "0:00",
    }


def compute_overtime_projection(metrics: Dict[str, str]) -> Dict[str, str]:
    early_total_min = time_str_to_minutes(metrics["早出残業合計"])
    work_days = float(metrics["出勤日数"])
    planned_days = float(metrics["所定日数"])
    holiday_days = float(metrics["年休日数"])

    remaining_days = (planned_days - work_days) + holiday_days
    avg_overtime_min = early_total_min / work_days if work_days else 0
    projected_total_min = avg_overtime_min * planned_days

    return {
        "平均残業時間_日": minutes_to_time_str(int(avg_overtime_min)),
        "残業予測_月末": minutes_to_time_str(int(projected_total_min)),
        "残り出勤数_日": f"{remaining_days:.1f}",
        "予測分_分": str(int(projected_total_min)),
    }


def show_overtime_alert_if_needed(projected_total_min: int) -> None:
    today = dt.datetime.today()
    if (today.day >= ZANGYOU_ALERT_DAY) and (
        projected_total_min >= ZANGYOU_LIMIT_HOUR * 60
    ):
        messagebox.showwarning(
            "⚠️ 残業時間注意",
            f"このままでは月末の残業時間が{ZANGYOU_LIMIT_HOUR}時間を超えます！\n予測: {minutes_to_time_str(projected_total_min)}",
        )
    else:
        log("残業アラートの条件には該当しません。")


# =============================================================================
# メイン処理
# =============================================================================


def main() -> None:
    drv: Optional[webdriver.Edge] = None
    try:
        password, reason = ask_password_and_reason()

        # 環境パス
        script_dir = Path(os.getcwd())
        excel_path = script_dir / EXCEL_FILENAME
        kv = _load_excel_kv(excel_path)

        login_id = _get_from_kv(kv, EXCEL_LOGIN_KEYS, required=True)

        # 定時（Excel優先、未設定なら既定17:00）
        fixed_off_text = _get_from_kv(kv, EXCEL_FIXED_OFF_KEYS, required=False)
        if fixed_off_text:
            try:
                hhmm = parse_hhmm(fixed_off_text)
                global FIXED_OFF_TIME
                FIXED_OFF_TIME = dt.datetime.strptime(hhmm, "%H:%M")
                log(f"Excel定義の定時を使用: {hhmm}")
            except Exception as e:
                warn(
                    f"定時の形式が不正です: {fixed_off_text} ({e})。既定17:00を使用します。"
                )
        else:
            log(
                f"Excelに定時未設定。既定{FIXED_OFF_TIME.strftime('%H:%M')}を使用します。"
            )

        driver_path = resolve_driver_path()
        drv = create_driver(driver_path)

        # ===== ログイン =====
        log("ログインページにアクセス中...")
        drv.get(LOGIN_URL)
        drv.find_element(By.NAME, "LoginID").send_keys(login_id)
        drv.find_element(By.NAME, "PassWord").send_keys(password)
        drv.find_element(By.NAME, "btnLogin").click()

        # ログイン直後のアラートを先に処理
        handle_possible_alert(drv, timeout=3)

        # frame待機
        wait(drv, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "frame")))

        # ===== 出勤/退勤クリック =====
        handle_possible_alert(drv, timeout=1)
        if not IS_TEST:
            clicked = find_and_click_in_frames(
                drv, By.LINK_TEXT, "退　勤", click=True, frame_wait=3
            )
        else:
            clicked = find_and_click_in_frames(
                drv, By.LINK_TEXT, "出　勤", click=True, frame_wait=3
            )
        if not clicked:
            warn("出勤/退勤リンクが見つかりませんでした。")
            return

        # ===== 打刻時刻取得（ポップアップ）=====
        main_window = drv.current_window_handle
        switch_to_new_window(drv, timeout=5)

        punch_time = get_punch_time_from_popup(drv)
        if not punch_time:
            warn("打刻時間が取得できませんでした。")
            return
        log(f"打刻時間: {punch_time}")

        # ポップアップ閉じ
        try:
            drv.find_element(By.LINK_TEXT, "戻る").click()
            log("ポップアップを閉じました。")
            wait(drv, 3).until(lambda d: len(d.window_handles) == 1)
        except Exception as e:
            warn(f"戻るボタン操作失敗: {e}")

        drv.switch_to.window(main_window)
        drv.switch_to.default_content()

        # ===== 残業申請（理由ありのとき）=====
        if reason is None:
            log("残業申請はスキップします。")
        else:
            start_hm = FIXED_OFF_TIME.strftime("%H:%M")
            end_hm = punch_time
            log(f"残業申請時間: {start_hm} ～ {end_hm}")

            navigate_menu_to_overtime_form(drv)
            fill_overtime_form(drv, start_hm, end_hm, reason)

        # ===== 申請直後にそのまま週報へ遷移して予測 =====
        log("残業時間予測のため週報へ遷移します。")
        navigate_to_weekly_report(drv)
        metrics = extract_weekly_metrics(drv)
        proj = compute_overtime_projection(metrics)

        print("======== [INFO] 残業予測モニタリング ========")
        print(f"・平均残業時間/日: {proj['平均残業時間_日']}")
        print(f"・残業時間予測（月末）: {proj['残業予測_月末']}")
        print(f"・月の残り出勤数: {proj['残り出勤数_日']} 日")
        print("===========================================")

        print("\n【📈 残業予測】")
        print(f"- 平均残業時間/日: {proj['平均残業時間_日']}")
        print(f"- 残業時間予測（月末）: {proj['残業予測_月末']}")
        print(f"- 月の残り出勤数: {proj['残り出勤数_日']} 日")

        show_overtime_alert_if_needed(int(proj["予測分_分"]))

        # ===== 5分待ってから自動終了 =====
        log(f"ブラウザを {BROWSER_AUTO_CLOSE_AFTER_SEC} 秒後に自動終了します。")
        time.sleep(BROWSER_AUTO_CLOSE_AFTER_SEC)

    except SystemExit:
        raise
    except WebDriverException as e:
        err(f"Seleniumエラー: {e}")
        sys.exit(1)
    except Exception as e:
        err(str(e))
        sys.exit(1)
    finally:
        try:
            if drv is not None:
                drv.quit()
                log("ブラウザを閉じました。")
        except Exception:
            pass


if __name__ == "__main__":
    main()
