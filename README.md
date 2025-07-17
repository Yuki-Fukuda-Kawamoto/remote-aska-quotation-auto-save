# 🧾 aska-quotation-auto-save

## 📌 プロジェクト概要

このプロジェクトは、川本製作所向けの見積書管理業務において、  
**開いているExcelファイル（見積書）から必要情報を抽出し、提出用フォルダに自動保存・整備** する処理をPythonで自動化します。

Power Automate Desktop で構築されていた手作業を Python に置き換えることで、  
作業効率向上・人的ミス削減・メンテナンス性の向上を図ります。

---

## 🎯 処理の目的

- 開いている見積書ファイル（`.xlsm`）から、
  - 見積番号（S1セル）
  - 顧客名（C1セル → ASC変換 → 英語化）
  - 案件名（C6セル）
  を自動で取得
- 上記情報から提出フォルダ名を生成し、所定の保存先に `.xlsm` ファイルを保存
- 保存先のパスをクリップボードにコピー（貼り付け用）

---

## 🗂️ フォルダ構成（推奨）

aska-quotation-auto-save/
├── modules/ # 各処理のPythonモジュール
│ ├── load_excel_data.py # Excelから情報を抽出する
│ ├── make_folder_name.py # 提出フォルダ名を組み立てる
│ ├── save_excel_to_folder.py # Excelファイルを所定フォルダに保存
│ └── clipboard_util.py # 保存先フォルダ名をクリップボードにコピー
├── temp_file/ # 一時ファイル置き場（.gitignore管理）
│ └── .gitkeep
├── logs/ # ログファイル保存先
│ └── .gitkeep
├── main.py # メイン実行ファイル
├── config.py # パスやファイル名などの設定
└── README.md # このファイル


---

## 🚀 処理の流れ

1. アクティブな見積書ファイル（例：`ASKA101P11.xlsm`）を取得
2. 指定セルの値（S1, C1, C6）を読み取る
3. 顧客名を英字に変換（ASC）し、1語目を抽出
4. 日付＋見積番号＋顧客名＋案件名を結合してフォルダ名を生成
5. 指定パスにフォルダを作成し、ファイルを `.xlsm` 形式で保存
6. フォルダ名をクリップボードにコピー

---

## 🛠 使用ライブラリ

- `xlwings`（Excelファイルの操作）
- `os`, `datetime`, `re`（標準ライブラリ）
- `pyperclip`（クリップボード操作）

---

## 🧭 今後の開発ロードマップ

| フェーズ   | 内容 |
|---------  |------|
| ✅ Step 1 | Git初期化・構成設計・README作成・config.py作成 |
※2025/07/11 完了

| 🚧 Step 2 | Excel情報取得モジュール（load_excel_data.py）実装 |
関数：extract_excel_info()
機能：開いているかつファイル名が「EXCEL_FILE_NAME = "ASKA101P11.xlsm"（config.pyで設定）」から情報を読み取る
引数：なし
戻り値：
return {
        "quotation_no": quotation_no,
        "quotation_project": quotation_project,
        "quotation_customer": quotation_customer,
        "customer_text": customer_text
    }
※2025/07/11 完了


| 🔜 Step 3 | フォルダ名生成ロジック（make_folder_name.py） |
関数：generate_folder_name(quotation_no: str, customer_text: str, project_name: str)
引数：extract_excel_info()で取得した値。
戻り値：保存日【Q】見積番号_案件名
例）20250717【Q】KMT-7619-032267-01_MICROTECHNICS SYSTEMS CO._Consumable Parts for GE-ME-type (with Past Sales)   

関数： generate_file_name(quotation_no: str)
引数：extract_excel_info()で取得した値。
戻り値：見積番号.xlsm
例：KMT-7619-032267-01.xlsm 

関数：enerate_full_path(folder_name: str, file_name: str) -> Path:
引数：extract_excel_info()で取得した値。
戻り値：QUOTATION_BASE_DIR\generate_folder_name（）で作成したディレクトリ名\ generate_file_nameで作成したファイル名
例） C:\Users\2012003\OneDrive - 株式会社川本製作所\2012003_福田\01_営業\10_見積もり\01_提出見積もり\20250717【Q】KMT-7619-032267-01_MICROTECHNICS SYSTEMS CO._Consumable Parts for GE-ME-type (with Past Sales)\KMT-7619-032267-01.xlsm

上記３つを統合する関数🧩 全てまとめて生成（主にmainから呼ばれる）
関数名：make_save_path(info: dict) -> dict:
引数：extract_excel_info()で取得した値。
戻り値：
return {
        "folder_name": folder_name,
        "file_name": file_name,
        "full_path": full_path
    }

| 🔜 Step 4 | 保存・コピー処理（save_excel_to_folder.py / clipboard_util.py） |
| 🔜 Step 5 | `main.py` 統合・ログ出力対応・例外処理追加 |
| ⏳ Step 6 | テスト用Excel準備・動作検証・業務反映 |

---

## 📬 補足

- 保存先のパスや命名ルールは `config.py` にて一元管理予定です。
- Power Automate Desktop で使用していたフローとほぼ同等の動作を Python で再現します。

---

## 🧑‍💻 開発者メモ

- 見積書ファイルは「すでに開いている状態」で使用する前提
- ファイル名は（例：`ASKA101P11.xlsm`）のように識別可能であること
- ASC関数の代替処理は `=ASC()` 相当をシートに書き込み → 結果読み取りで対応

