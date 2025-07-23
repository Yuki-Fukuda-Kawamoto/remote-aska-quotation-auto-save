# 📌 ASKA見積書 保存パス生成メインスクリプト
# --------------------------------------------------
# ・対象Excelから見積情報を抽出
# ・保存用のフォルダ名・ファイル名・フルパスを生成
# ・結果を画面に表示
# --------------------------------------------------

import logging
from tabulate import tabulate

# 🧱 初期設定
from config import (
    LOGGING_FORMAT, 
    EXCEL_FILE_NAME, 
    DATE_FORMAT, 
    MAX_PROJECT_NAME_LENGTH, 
    QUOTATION_BASE_DIR)

logging.basicConfig(level=logging.INFO, format=LOGGING_FORMAT)

# 📩 見積情報取得モジュール
from modules.load_askaquotation_excel import extract_excel_info

# 📁 保存パス生成モジュール
from modules.generate_save_path import make_save_path

# ✅ メイン処理
def main():
    logging.info("🚀 見積情報の取得と保存パス生成を開始します")

    try:
        # ① Excelから情報取得
        info = extract_excel_info(EXCEL_FILE_NAME)
        logging.info("📥 見積情報の読み取りに成功しました")

        # 取得した情報を表示
        logging.info("📋 見積情報を表示します")
        logging.info(tabulate(info.items(), headers=["項目", "値"], tablefmt="grid"))   
        info.update({
            'DATE_FORMAT' : DATE_FORMAT,
            "MAX_PROJECT_NAME_LENGTH" : MAX_PROJECT_NAME_LENGTH,
            "QUOTATION_BASE_DIR": QUOTATION_BASE_DIR
        })
            
        # ② 保存パスの生成
        path_info = make_save_path(info)
        logging.info("📦 保存パス情報の生成に成功しました")

        # ③ 結果の表示
        print("\n📋 保存先情報（確認用）")
        print(tabulate(path_info.items(), headers=["項目", "値"], tablefmt="grid"))

    except Exception as e:
        logging.error(f"❌ エラーが発生しました: {e}")

# ✅ 単体実行
if __name__ == "__main__":
    main()
