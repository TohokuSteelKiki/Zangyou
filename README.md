# Auto残業申請Bot

## 🧰 主な機能

下記の様に退勤から残業申請の入力までする
実行 → エクセルからID取得 → パスワード入力ポップ →申請有無確認ポップ→  申請理由入力ポップ→ログイン
→退勤（テストのため出勤をクリック）
→ 退勤時間取得（ポップから取得PC時計と時差があるため）

申請しない場合
→終了

申請する場合
残業申請へ移行 → 残業申請入力 → 登録

---

## 📁 ディレクトリ構成

auto_zangyo/
├── zangyo_script.py # メインスクリプト
├── IDPASS.xlsx # 最低限IDを入力
├── README.md #

---

## 📋 事前準備

### 0．仮想環境の実装

```bash
python -m venv .venv
```

作成した仮想環境を有効化します。

- Windows:

    ```ps
    .venv\Scripts\activate
    ```

- macOS/Linux:

    ```ps
    .venv\Include\activate   
    ```

### A.1. Pythonパッケージのインストール selenium版

```ps
python.exe -m pip install --upgrade pip
pip install selenium webdriver-manager pandas tkinter
```

### B.1. Pythonパッケージのインストール pywinauto版

```ps
python.exe -m pip install --upgrade pip
pip install pywinauto pandas openpyxl pyperclip tkinter
```

### 2. IDPASS.xlsxの構成

スクリプトID
TimeProGX your_id

「スクリプト」列にはTimeProGXを指定

「ID」列にログインIDを記載

🧱 必要なインストール項目一覧

| 区分         | 名称  | インストール方法・備考 |
|--------------|-----------------|-----------|
| Python本体   | Python 3.7以降  | [公式サイト](https://www.python.org/) よりダウンロード |
| Pythonパッケージ | selenium | `pip install selenium` |
| | pandas | `pip install pandas` |
| | openpyxl | `pip install openpyxl`（Excelファイル読込に必要） |
| Webブラウザ | Google Chrome   | 最新版推奨。自動化対象として使用 |
| GUIモジュール | tkinter（標準） | Pythonに標準同梱（ない場合は `sudo apt install python3-tk` などで追加） |
| Excelファイル | IDPASS.xlsx | スクリプトと同じディレクトリに配置（ID情報を記載） |

※ pipが使えない場合はpython -m pip installパッケージ名を使用してください。

## 🚀 実行方法

python main.py
起動後の操作
ログインパスワード（非表示）の入力
残業理由（例：「納期対応」など）の入力

✅ 残業時間の扱い
定時（17:00）からの退勤までの時間（最小1分単位）
残業時間乖離が何分からなのか不明のため現状維持でする→17：10以降は必ず入力する
17:00～17:10の間に関しては申請の有無を申請者に委ねる。
