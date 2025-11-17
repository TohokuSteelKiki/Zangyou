# Auto残業申請Bot
<<<<<<< HEAD
ソフトバージョン: V1.0.4
webdriver：  142.0.3595.80
=======
ソフトバージョン: V1.0.3
webdriver： 142.0.3595.53
>>>>>>> 93fb1d10763b02954df0f86c18ee3334744ff7aa
## 🧰 主な機能

下記の様に退勤から残業申請の入力までする
実行 → エクセルからID取得 → パスワード入力ポップ →申請有無確認ポップ→  申請理由入力ポップ→ログイン
→退勤（テストのため出勤をクリック）
→ 退勤時間取得（ポップから取得PC時計と時差があるため）

申請しない場合
→終了

申請する場合
残業申請へ移行 → 残業申請入力 → 登録
→ 就業情報 → 就業日次処理 → 就業週報→ 月末残業時間の予測
→所定日数/出勤日数/年休日数/早出残業合計からデータ取得
→ 警告：当月20日以降 かつ 予測≥30h で警告ダイアログ
→ 300秒待機 → ブラウザ自動終了


---

## 📁 ディレクトリ構成

auto_zangyo/
├── main.py # メインスクリプト
├── test.py # テストスクリプト
├── ID.xlsx # 最低限IDを入力
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
pip install selenium webdriver-manager pandas openpyxl tkinter
```

### B.1. Pythonパッケージのインストール pywinauto版

```ps
python.exe -m pip install --upgrade pip
pip install pywinauto pandas openpyxl pyperclip tkinter
```

### 2. ID.xlsxの構成

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
| | tkinter | `pip install tkinter`PythonでのGUI表現として使用 |
| | selenium | `pip install selenium`Pythonでのブラウザ自動化として使用 |
| | webdriver-manager | `pip install webdriver-manager`Pythonでのブラウザ自動化として使用 |
| | pywinauto | `pip install pywinauto`Pythonでのブラウザ自動化として使用 |
| | pyperclip | `pip install pyperclip`Pythonでクリップボードとして使用 |
| Webブラウザ | Microsoft Edge   | 最新版推奨。自動化対象として使用 |
| GUIモジュール | tkinter（標準） | Pythonに標準同梱（ない場合は `sudo apt install python3-tk` などで追加） |
| Excelファイル | IDPASS.xlsx | スクリプトと同じディレクトリに配置（ID情報を記載） |

※ pipが使えない場合はpython -m pip installパッケージ名を使用してください。

## 🚀 実行方法

main.py
起動後の操作
ログインパスワード（非表示）の入力
残業理由（例：「納期対応」など）の入力

✅ 残業時間の扱い
定時（17:00）からの退勤までの時間（最小1分単位）
残業時間乖離が何分からなのか不明のため現状維持でする→17：10以降は必ず入力する
17:00～17:10の間に関しては申請の有無を申請者に委ねる。

## 📦 配布について

main.py を pyInstaller で exe 化する。(アイコンも含む)
コマンド （ターミナルで実行）
python -m PyInstaller --onefile --icon=icon.ico main.py

実行後に dist フォルダが生成される。

python -m PyInstaller  main.spec

ID.xlsx と msedgedriver.exe を同梱する。
※msedgedriver.exeはEdgeのWEBドライバーで下記からダウンロード
https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver?form=MA13LH&ch=1#downloads



配布時ディレクトリ構成
dist/  (フォルダ名は変更可)
├── ID.xlsx          # 名前変更不可
├── main.exe         # 名前変更可
├── msedgedriver.exe # 名前変更不可
├── test.exe 
├── 残業申請_半自動化ツール取扱説明書.pdf


※ フォルダ内ファイルの移動は禁止。

## 配布後
1. ID.xlsxのID欄に社員コード（４桁）を記入
2. main.exe をクリックしスクリプトを実行サせる


## トラブルシュート
Edgeをアップデートなどをしてから
mai.exeを実行した際にログイン画面（ネット接続）で失敗した場合
Edge本体とWEBドライバーのバージョンによる互換性がなくTimeProを開けないときがある。
対処
dist内のmsedgedriver.exe を対応するモノに更新する
