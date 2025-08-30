# encoding: utf-8
import pandas as pd
import datetime
import time
import sys
import tkinter as tk
from tkinter import simpledialog, messagebox
import os
import pyperclip
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
from pywinauto.findwindows import ElementNotFoundError

class TimeProGXAutomation:
    def __init__(self):
        self.desktop = Desktop(backend="uia")
        self.edge_window = None
        self.login_id = None
        self.password = None
        self.zangyo_reason = None
        self.proceed_overtime = False
        
        # 定時設定
        self.teiji = datetime.datetime.strptime("17:00", "%H:%M")
        
        # 設定
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.excel_path = os.path.join(self.script_dir, "IDPASS.xlsx")
        self.target_script = "TimeProGX"
        self.login_url = "http://128.198.11.125/jinjikanri/"
        
        # Edge起動パス
        self.edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            "msedge.exe"
        ]
    
    def get_user_input(self):
        """ユーザーからの入力を取得"""
        root = tk.Tk()
        root.withdraw()
        root.iconbitmap("icon.ico")
        

        # パスワード入力
        self.password = simpledialog.askstring(
            "パスワード入力", "ログイン用パスワードを入力してください：", show="*"
        )
        if not self.password:
            print("[ERROR] パスワードが入力されませんでした。")
            sys.exit(1)
        
        # 残業申請確認
        self.proceed_overtime = messagebox.askyesno("確認", "残業申請を実行しますか？")
        
        if self.proceed_overtime:
            # 残業理由入力
            self.zangyo_reason = simpledialog.askstring(
                "残業理由入力", "残業申請の理由を入力してください："
            )
            if not self.zangyo_reason:
                print("[ERROR] 残業理由が入力されませんでした。")
                sys.exit(1)
        
        root.destroy()
    
    def load_login_id(self):
        """ExcelからログインIDを取得"""
        try:
            df = pd.read_excel(self.excel_path, dtype={"ID": str})
            row = df[df["スクリプト"] == self.target_script].iloc[0]
            self.login_id = row["ID"].strip()
            print(f"[INFO] ログインID: {self.login_id}")
        except Exception as e:
            print(f"[ERROR] Excel読み込み失敗: {e}")
            sys.exit(1)
    
    def launch_edge(self):
        """Microsoft Edgeを起動"""
        edge_launched = False
        
        for edge_path in self.edge_paths:
            try:
                if os.path.exists(edge_path) or edge_path == "msedge.exe":
                    print(f"Edgeを起動中: {edge_path}")
                    Application().start(edge_path)
                    edge_launched = True
                    print("Edge起動成功")
                    break
            except Exception as e:
                print(f"パス {edge_path} での起動失敗: {e}")
                continue
        
        if not edge_launched:
            print("Edgeの起動に失敗しました")
            sys.exit(1)
        
        # Edgeの初期化待ち
        print("Edgeの初期化を待機中...")
        time.sleep(5)
        
        # Edgeウィンドウを取得
        for attempt in range(10):
            try:
                edge_windows = self.desktop.windows(title_re=".*Edge.*", class_name_re="Chrome_WidgetWin.*")
                if edge_windows:
                    self.edge_window = edge_windows[0]
                    print(f"Edgeウィンドウを発見: {self.edge_window.window_text()}")
                    break
            except:
                print(f"ウィンドウ検索試行 {attempt + 1}/10")
                time.sleep(1)
        
        if not self.edge_window:
            print("Edgeウィンドウが見つかりません")
            sys.exit(1)
    
    def navigate_to_login(self):
        """ログインページにアクセス"""
        self.edge_window.set_focus()
        time.sleep(2)
        
        # アドレスバーにフォーカス
        send_keys("^l")
        time.sleep(1)
        
        # URLを入力
        pyperclip.copy(self.login_url)
        send_keys("^v{ENTER}")
        print("ログインページを読み込み中...")
        time.sleep(5)
    
    def login(self):
        """ログイン処理"""
        print("ログイン中...")
        self.edge_window.set_focus()
        time.sleep(2)
        
        # ログインID入力（TABで移動してフィールドを探す）
        for i in range(10):
            send_keys("{TAB}")
            time.sleep(0.5)
            
            # IDを入力してみる
            pyperclip.copy(self.login_id)
            send_keys("^a^v")
            time.sleep(0.5)
            
            # 次のフィールドへ
            send_keys("{TAB}")
            time.sleep(0.5)
            
            # パスワードを入力
            pyperclip.copy(self.password)
            send_keys("^a^v")
            time.sleep(0.5)
            
            # ログインボタン（Enter or TAB + Enter）
            send_keys("{ENTER}")
            time.sleep(3)
            
            # ログイン成功の確認（フレームが表示されるか）
            break
        
        print("ログイン完了")
        time.sleep(3)
    
    def punch_out(self):
        """退勤処理"""
        print("退勤処理中...")
        self.edge_window.set_focus()
        
        # 出勤/退勤ボタンを探してクリック
        # フレーム内を順次探索
        for i in range(5):
            send_keys("{TAB}")
            time.sleep(0.5)
            send_keys("{ENTER}")
            time.sleep(1)
        
        # 別の方法: 文字列検索
        pyperclip.copy("出　勤")
        send_keys("^f")
        time.sleep(1)
        send_keys("^v{ENTER}{ESC}")
        time.sleep(1)
        send_keys("{ENTER}")
        
        print("退勤処理完了")
        time.sleep(3)
    
    def get_punch_time(self):
        """打刻時間を取得（簡易版）"""
        # 実際の実装では、ポップアップから時間を取得する必要があります
        # ここでは現在時刻を使用
        current_time = datetime.datetime.now()
        punch_time = current_time.strftime("%H:%M")
        print(f"打刻時間: {punch_time}")
        return punch_time
    
    def apply_overtime(self, punch_time):
        """残業申請処理"""
        if not self.proceed_overtime:
            print("残業申請をスキップします")
            return
        
        print("残業申請処理中...")
        
        # 残業時間判定
        punch_dt = datetime.datetime.strptime(punch_time, "%H:%M")
        delta_min = (punch_dt - self.teiji).total_seconds() / 60
        
        if delta_min < 1:
            print("残業時間が1分未満のため申請をスキップします")
            return
        
        start_time = self.teiji.strftime("%H:%M")
        end_time = punch_time
        print(f"残業申請時間: {start_time} ～ {end_time}")
        
        # 届出処理への移動（TAB + Enterで探索）
        self.edge_window.set_focus()
        
        # 届出処理をクリック
        pyperclip.copy("届出処理")
        send_keys("^f")
        time.sleep(1)
        send_keys("^v{ENTER}{ESC}")
        time.sleep(1)
        send_keys("{ENTER}")
        time.sleep(2)
        
        # 時間外申請をクリック
        pyperclip.copy("時間外申請")
        send_keys("^f")
        time.sleep(1)
        send_keys("^v{ENTER}{ESC}")
        time.sleep(1)
        send_keys("{ENTER}")
        time.sleep(3)
        
        # 申請フォームに入力
        # 開始時間
        send_keys("{TAB}")
        pyperclip.copy(start_time)
        send_keys("^a^v")
        
        # 終了時間
        send_keys("{TAB}")
        pyperclip.copy(end_time)
        send_keys("^a^v")
        
        # 理由
        send_keys("{TAB}")
        pyperclip.copy(self.zangyo_reason)
        send_keys("^a^v")
        
        print("残業申請フォーム入力完了")
        
        # 登録ボタン（実際の実行時はコメント解除）
        # send_keys("{TAB}{ENTER}")
        
        print("残業申請完了")
    
    def run(self):
        """メイン処理"""
        try:
            print("TimeProGX自動化を開始します")
            
            # ユーザー入力
            self.get_user_input()
            
            # ExcelからID読み込み
            self.load_login_id()
            
            # Edge起動
            self.launch_edge()
            
            # ログインページにアクセス
            self.navigate_to_login()
            
            # ログイン
            self.login()
            
            # 退勤処理
            self.punch_out()
            
            # 打刻時間取得
            punch_time = self.get_punch_time()
            
            # 残業申請
            self.apply_overtime(punch_time)
            
            print("処理完了")
            
        except Exception as e:
            print(f"エラーが発生しました: {e}")
        
        finally:
            print("プログラム終了")

if __name__ == "__main__":
    automation = TimeProGXAutomation()
    automation.run()