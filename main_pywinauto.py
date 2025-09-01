# encoding: utf-8
import pandas as pd
import datetime
import time
import sys
import tkinter as tk
from tkinter import simpledialog, messagebox
import os
import re
import pyperclip
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
from pywinauto.findwindows import ElementNotFoundError


class TimeProGXAutomation:
    def __init__(self):
        self.desktop = Desktop(backend="uia")
        self.edge_window = None
        self.login_id = ""
        self.password = ""
        self.zangyo_reason = None
        self.proceed_overtime = False

        # 定時
        self.teiji = datetime.datetime.strptime("17:00", "%H:%M")

        # 設定
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.excel_path = os.path.join(self.script_dir, "IDPASS.xlsx")
        self.target_script = "TimeProGX"
        self.login_url = "http://128.198.11.125/xgweb/frame.asp"

        # Edge起動パス候補
        self.edge_paths = [
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            "msedge.exe",
        ]

    # ========== 小物ユーティリティ ==========
    def _paste(self, text: str):
        pyperclip.copy(text)
        send_keys("^v")

    def _wait_win(self, title_re, timeout=20):
        end = time.time() + timeout
        win = None
        while time.time() < end:
            try:
                win = self.desktop.window(
                    title_re=title_re, class_name_re="Chrome_WidgetWin.*"
                )
                if win.exists() and win.is_visible():
                    return win
            except Exception:
                pass
            time.sleep(0.3)
        raise ElementNotFoundError(f"window not found: {title_re}")

    def _try_child(self, parent, **kwargs):
        try:
            c = parent.child_window(**kwargs)
            if c.exists():
                return c
        except Exception:
            pass
        return None

    def _find_click_text(
        self, parent, titles, control_types=("Hyperlink", "Text"), timeout=5
    ):
        end = time.time() + timeout
        while time.time() < end:
            for t in titles:
                for ct in control_types:
                    ctrl = self._try_child(parent, title=t, control_type=ct)
                    if ctrl:
                        try:
                            ctrl.click_input()
                            return True
                        except Exception:
                            pass
            time.sleep(0.2)
        return False

    def _search_and_enter(self, keyword):
        # Ctrl+Fで検索→Enterでフォーカス→Enterで起動
        send_keys("^f")
        time.sleep(0.3)
        pyperclip.copy(keyword)
        send_keys("^v{ENTER}{ESC}")
        time.sleep(0.3)
        send_keys("{ENTER}")

    # ========== 入力取得 ==========
    def get_user_input(self):
        root = tk.Tk()
        root.withdraw()
        try:
            try:
                root.iconbitmap("icon.ico")
            except Exception:
                pass

            self.password = simpledialog.askstring(
                "パスワード入力", "ログイン用パスワードを入力してください：", show="*"
            )
            if not self.password:
                print("[ERROR] パスワード未入力")
                sys.exit(1)

            self.proceed_overtime = messagebox.askyesno(
                "確認", "残業申請を実行しますか？"
            )

            if self.proceed_overtime:
                self.zangyo_reason = simpledialog.askstring(
                    "残業理由入力", "残業申請の理由を入力してください："
                )
                if not self.zangyo_reason:
                    print("[ERROR] 残業理由未入力")
                    sys.exit(1)
        finally:
            root.destroy()

    # ========== ExcelからID ==========
    def load_login_id(self):
        try:
            df = pd.read_excel(self.excel_path, dtype={"ID": str})
            row = df[df["スクリプト"] == self.target_script].iloc[0]
            self.login_id = row["ID"].strip()
            print(f"[INFO] ログインID: {self.login_id}")
        except Exception as e:
            print(f"[ERROR] Excel読み込み失敗: {e}")
            sys.exit(1)

    # ========== Edge起動 ==========
    def launch_edge(self):
        launched = False
        for p in self.edge_paths:
            try:
                if p == "msedge.exe" or os.path.exists(p):
                    print(f"[INFO] Edge起動: {p}")
                    Application().start(p)
                    launched = True
                    break
            except Exception as e:
                print(f"[WARN] 起動失敗 {p}: {e}")

        if not launched:
            print("[ERROR] Edgeの起動に失敗")
            sys.exit(1)

        print("[INFO] Edge初期化待機")
        time.sleep(2)
        # 何かしらのEdgeトップウィンドウを捕まえる
        self.edge_window = self._wait_win(r".*Edge.*", timeout=20)
        print(f"[INFO] Edgeウィンドウ: {self.edge_window.window_text()}")

    # ========== ログインページへ ==========
    def navigate_to_login(self):
        self.edge_window.set_focus()
        time.sleep(0.5)
        send_keys("^l")
        time.sleep(0.3)
        pyperclip.copy(self.login_url)
        send_keys("^v{ENTER}")
        print("[INFO] ログインページ遷移")
        time.sleep(3)

    # ========== ログイン ==========
    def login(self):
        print("[INFO] ログイン開始")
        # コンテンツ領域にEditが出るまで試す → 失敗時はTABフォールバック
        win = self._wait_win(r".*TimePro.*Edge.*", timeout=20)
        win.set_focus()
        time.sleep(0.5)

        # 直接Editを探す
        edits = []
        try:
            edits = win.descendants(control_type="Edit")
        except Exception:
            edits = []

        if len(edits) >= 2:
            try:
                edits[0].click_input()
                self._paste(self.login_id)
                time.sleep(0.2)
                edits[1].click_input()
                self._paste(self.password)
                time.sleep(0.2)
                send_keys("{ENTER}")
            except Exception:
                # フォールバック
                self._tab_login_fallback(win)
        else:
            self._tab_login_fallback(win)

        time.sleep(3)
        print("[INFO] ログイン完了想定")

    def _tab_login_fallback(self, win):
        # TABでID→PW→Enter
        win.set_focus()
        for _ in range(10):
            send_keys("{TAB}")
            time.sleep(0.15)
            self._paste(self.login_id)
            time.sleep(0.15)
            send_keys("{TAB}")
            time.sleep(0.15)
            self._paste(self.password)
            time.sleep(0.15)
            send_keys("{ENTER}")
            time.sleep(2)
            break

    # ========== 退勤打刻 ==========
    def punch_out(self):
        print("[INFO] 退勤打刻")
        top = self._wait_win(r".*TimePro.*Edge.*", timeout=25)
        top.set_focus()
        time.sleep(0.5)

        # まずはUI要素直接クリックを試す
        clicked = self._find_click_text(
            top,
            titles=("退　勤", "退勤"),
            control_types=("Hyperlink", "Text"),
            # top, titles=("退　勤", "退勤"), control_types=("Hyperlink", "Text")
        )
        if not clicked:
            # 見つからなければ検索フォールバック

            self._search_and_enter("" "")
            # self._search_and_enter("退　勤")

        # 打刻結果ダイアログを待ち、打刻時刻を取得して閉じる
        punch_time = self._handle_punch_result()
        print(f"[INFO] 打刻時刻: {punch_time}")
        return punch_time

    def _handle_punch_result(self):
        # 打刻結果ウィンドウが出るまで待機し、時刻を抽出 → 「閉じる」
        t0 = time.time()
        dlg = None
        while time.time() - t0 < 10:
            try:
                dlg = self.desktop.window(title_re=r"打刻結果.*Edge.*")
                if dlg.exists() and dlg.is_visible():
                    break
            except Exception:
                pass
            time.sleep(0.5)

        punch_time = None
        if dlg and dlg.exists():
            try:
                # テキストを総なめして時刻を抽出
                texts = []
                for c in dlg.descendants():
                    try:
                        txt = c.window_text()
                        if txt:
                            texts.append(txt)
                    except Exception:
                        pass
                joined = "\n".join(texts)
                m = re.search(r"([01]?\d|2[0-3]):[0-5]\d", joined)
                if m:
                    punch_time = m.group(0)
            except Exception:
                pass

            # 閉じる
            closed = self._find_click_text(
                dlg, titles=("閉じる",), control_types=("Button",)
            )
            if not closed:
                # キーで閉じる
                send_keys("{ESC}")
        else:
            print("[WARN] 打刻結果ダイアログ未検出。システム時刻を使用")

        if not punch_time:
            punch_time = datetime.datetime.now().strftime("%H:%M")
        return punch_time

    # ========== 残業申請 ==========
    def apply_overtime(self, punch_time):
        if not self.proceed_overtime:
            print("[INFO] 残業申請スキップ")
            return

        # 残業時間判定
        punch_dt = datetime.datetime.strptime(punch_time, "%H:%M")
        delta_min = (punch_dt - self.teiji).total_seconds() / 60.0
        if delta_min < 10:
            print("[INFO] 残業<1分。申請スキップ")
            return

        start_time = self.teiji.strftime("%H:%M")
        end_time = punch_time
        print(f"[INFO] 残業: {start_time} - {end_time}")

        main = self._wait_win(r".*TimePro.*Edge.*", timeout=20)
        main.set_focus()
        time.sleep(1.0)

        # 届出処理 → 就業届出処理 → 時間外申請
        if not self._find_click_text(main, ("届出処理",), ("Hyperlink", "Text")):
            self._search_and_enter("届出処理")
            time.sleep(1.0)

        main = self._wait_win(r".*TimePro.*Edge.*", timeout=10)
        if not self._find_click_text(main, ("就業届出処理",), ("Hyperlink", "Text")):
            self._search_and_enter("就業届出処理")
            time.sleep(1.0)

        main = self._wait_win(r".*TimePro.*Edge.*", timeout=10)
        if not self._find_click_text(main, ("時間外申請",), ("Hyperlink", "Text")):
            self._search_and_enter("時間外申請")
            time.sleep(1.5)

        # 入力フォーム操作
        form = self._wait_win(r".*TimePro.*Edge.*", timeout=10)
        form.set_focus()
        time.sleep(0.3)

        # 目安: Recorderの Edit#[1,0] → 開始、次のEdit → 終了、CheckBox#[0,1]、Edit#[6,0] → 理由
        edits = []
        checks = []
        try:
            edits = form.descendants(control_type="Edit")
            checks = form.descendants(control_type="CheckBox")
        except Exception:
            pass

        def _safe_set(edit_ctrl, value):
            try:
                edit_ctrl.click_input()
                send_keys("^a")
                self._paste(value)
                time.sleep(0.1)
                return True
            except Exception:
                return False

        # インデックス推定（存在チェックしつつRecorder準拠を優先）
        start_idx = 1 if len(edits) > 1 else 0
        end_idx = start_idx + 1 if len(edits) > start_idx + 1 else start_idx
        reason_idx = 11 if len(edits) > 11 else (len(edits) - 1 if edits else 0)
        check_idx = 1 if len(checks) > 1 else (0 if checks else None)

        ok = 0

        # 開始 17:00
        if edits:
            if _safe_set(edits[start_idx], start_time):
                ok += 1
        else:
            # フォールバック: TABで移動
            send_keys("{TAB}")
            self._paste(start_time)

        # 終了 打刻時刻
        if edits and end_idx != start_idx:
            # 多くの画面は4桁入力で自動次フィールドへ進むが、TABで確実に。
            try:
                edits[end_idx].click_input()
            except Exception:
                send_keys("{TAB}")
            if _safe_set(edits[end_idx], end_time):
                ok += 1
        else:
            send_keys("{TAB}")
            self._paste(end_time)

        # チェック ON
        if check_idx is not None:
            try:
                checks[check_idx].click_input()
            except Exception:
                pass

        # 理由
        if edits:
            if _safe_set(edits[reason_idx], self.zangyo_reason):
                ok += 1
        else:
            send_keys("{TAB}")
            self._paste(self.zangyo_reason)

        # 低信頼時はTABフォールバックで再入力
        if ok < 3:
            send_keys("{HOME}")  # フォーカス固定回避用（任意）
            send_keys("{TAB}")
            self._paste(start_time)
            send_keys("{TAB}")
            self._paste(end_time)
            send_keys("{TAB}")
            self._paste(self.zangyo_reason)

        print("[INFO] 残業申請フォーム入力完了")

        
        # 登録 or 照会→登録 TODO:登録の実装
        # self._find_click_text(form, ("照会",), ("Button", "Hyperlink", "Text"))
        # self._find_click_text(form, ("登録",), ("Button", "Hyperlink", "Text"))
        # if not self._find_click_text(form, ("登録",), ("Button", "Hyperlink", "Text")):
        #     self._find_click_text(form, ("照会",), ("Button", "Hyperlink", "Text"))
        #     time.sleep(1.0)
        #     self._find_click_text(form, ("登録",), ("Button", "Hyperlink", "Text"))

        time.sleep(1.5)
        print("[INFO] 残業申請完了想定")

    # ========== メイン ==========
    def run(self):
        try:
            print("[INFO] TimeProGX自動化開始")
            self.get_user_input()
            self.load_login_id()
            self.launch_edge()
            self.navigate_to_login()
            self.login()
            punch_time = self.punch_out()
            self.apply_overtime(punch_time)
            print("[INFO] 処理完了")
        except Exception as e:
            print(f"[ERROR] 例外: {e}")
        finally:
            print("[INFO] プログラム終了")


if __name__ == "__main__":
    automation = TimeProGXAutomation()
    automation.run()
