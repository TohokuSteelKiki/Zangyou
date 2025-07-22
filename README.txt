# ⏰ TimeProGX 自動退勤 & 残業申請スクリプト

このスクリプトは、**TimeProGX 勤怠管理システム**に自動ログインし、
「退勤ボタンの押下」および「残業申請（10分以上）」を自動実行するツールです。
ログインIDはExcelから取得、Seleniumを用いてブラウザ操作を自動化します。

---

## 🧰 主な機能

- ログインIDをExcel（IDPASS.xlsx）から取得
- GUIでパスワードおよび残業理由をユーザー入力
- TimeProGXにログインし、「退勤」リンクを自動クリック
- 打刻された退勤時刻を取得
- 残業時間を計算し、**10分未満はスキップ**
- 10分以上の場合、残業申請フォームに自動入力・送信

---

## 🖥️ 対応環境

- Windows OS（`tkinter` GUI使用）
- Python 3.7 以降
- Google Chrome（推奨）
- ChromeDriver（バージョン一致が必要）

---

## 📁 ディレクトリ構成

auto_zangyo/
├── zangyo_script.py # メインスクリプト
├── IDPASS.xlsx # スクリプトと同じフォルダに配置
├── README.md # 本ファイル


---

## 📋 事前準備

### 1. Pythonパッケージのインストール

```bash
pip install selenium pandas openpyxl

2. ChromeDriver のセットアップ
Chrome のバージョンと一致する chromedriver.exe をダウンロード

https://chromedriver.chromium.org/downloads

スクリプトと同じフォルダ、もしくは PATH に追加

3. IDPASS.xlsx の構成
スクリプト	ID
TimeProGX	your_id

「スクリプト」列には TimeProGX を指定

「ID」列にログインIDを記載

🧱 必要なインストール項目一覧
区分	名称	インストール方法・備考
Python本体	Python 3.7以降	https://www.python.org/ よりダウンロード
Pythonパッケージ	selenium	pip install selenium
pandas	pip install pandas
openpyxl	pip install openpyxl （Excelファイル読込に必要）
Webブラウザ	Google Chrome（最新版推奨）	自動化対象として使用
ドライバ	ChromeDriver	https://chromedriver.chromium.org/downloads にて、Chromeのバージョンに対応したものを入手
GUIモジュール	tkinter（標準）	Pythonに標準同梱（無い場合は sudo apt install python3-tk などで追加）
Excelファイル	IDPASS.xlsx	スクリプトと同じフォルダに配置（ID情報を記載）

※ pip が使えない場合は python -m pip install パッケージ名 を使用してください。

🚀 実行方法

python main.py
起動後の操作
ログインパスワード（非表示）の入力
残業理由（例：「納期対応」など）の入力

✅ 残業時間の扱い
定時（17:00）からの差分をもとに残業時間を判定
10分未満の残業は申請されません（自動スキップ）
10分以上の場合、開始時刻を「17:00」として申請が作成されます

