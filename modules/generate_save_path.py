# 📌 保存ファイル構造生成モジュール
# --------------------------------------------------
# 引数として受け取った見積情報（quotation_no, project_name, customer_text）から、
# フォルダ名・ファイル名・保存先フルパスを生成する。
# フォーマットに従って日付・記号・禁止文字処理・文字数制限なども考慮する。
# 処理の進行は logging にて視覚的に出力。
# --------------------------------------------------

# 📌 保存ファイル構造生成モジュール
# --------------------------------------------------
# 引数として受け取った見積情報から、
# ・フォルダ名
# ・ファイル名（.xlsm）
# ・フル保存パス（Path型）
# を生成して dict で返す。
# --------------------------------------------------

import re
import logging
from datetime import datetime
from pathlib import Path

# 📁 フォルダ名を生成
def generate_folder_name(quotation_no: str, customer_text: str, project_name: str) -> str:
    logging.info("📁 フォルダ名を生成中...")
    date_str = datetime.now().strftime(DATE_FORMAT)

    # 禁止文字の除去・置換
    safe_project = re.sub(r'[\\/:*?"<>|]', '_', project_name)

    # 文字数制限（切り詰め）
    if len(safe_project) > MAX_PROJECT_NAME_LENGTH:
        logging.info("✂️ 案件名が長すぎるためカットします")
        safe_project = safe_project[:MAX_PROJECT_NAME_LENGTH].rstrip()

    folder_name = f"{date_str}【Q】{quotation_no}_{customer_text}_{safe_project}"
    logging.info(f"✅ フォルダ名生成完了: {folder_name}")
    return folder_name

# 📄 ファイル名を生成
def generate_file_name(quotation_no: str) -> str:
    logging.info("📄 ファイル名を生成中...")
    file_name = f"{quotation_no}.xlsm"
    logging.info(f"✅ ファイル名生成完了: {file_name}")
    return file_name

# 📂 フル保存パスを生成
def generate_full_path(folder_name: str, file_name: str) -> Path:
    logging.info("📂 保存先パスを生成中...")
    full_path = QUOTATION_BASE_DIR / folder_name / file_name
    logging.info(f"✅ 保存パス生成完了: {full_path}")
    return full_path

# 🧩 全てまとめて生成（主にmainから呼ばれる）
def make_save_path(info: dict) -> dict:
    logging.info("🧲 保存パス情報一式を生成します")
    folder_name = generate_folder_name(
        quotation_no=info["quotation_no"],
        customer_text=info["customer_text"],
        project_name=info["quotation_project"]
    )
    file_name = generate_file_name(info["quotation_no"])
    full_path = generate_full_path(folder_name, file_name)

    return {
        "folder_name": folder_name,
        "file_name": file_name,
        "full_path": full_path
    }

# ✅ 動作確認用（直接実行時）
if __name__ == "__main__":
    import sys
    from pathlib import Path

    # 📁 config.py のある親フォルダをパスに追加
    sys.path.append(str(Path(__file__).resolve().parent.parent))

    # 🔁 再インポート（__main__ 限定）      
    from config import QUOTATION_BASE_DIR, DATE_FORMAT, MAX_PROJECT_NAME_LENGTH


    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    sample_info = {
        'quotation_no': 'KMT-7619-032267-01',
        'quotation_project': 'Consumable Parts for GE-ME-type (with Past Sales)',
        'quotation_customer': 'ＭＩＣＲＯＴＥＣＨＮＩＣＳ　ＳＹＳＴＥＭＳ　ＣＯ． ',
        'customer_text': 'MICROTECHNICS SYSTEMS CO.'
    }

    result = make_save_path(sample_info)

    from tabulate import tabulate
    print("\n📦 保存パス情報（確認用）")
    print(tabulate(result.items(), headers=["項目", "値"], tablefmt="grid"))
