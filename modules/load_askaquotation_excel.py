# 📩 load_askaquotation_excel.py - Excelファイルから必要情報を抽出する

import xlwings as xw
import logging
from tabulate import tabulate

def extract_excel_info() -> dict:
    logging.info("🧲 Excelファイルに接続中...")

    try:
        wb = xw.Book(EXCEL_FILE_NAME)
    except Exception as e:
        logging.error("⚠️ Excelファイルが開かれていないか、指定名が一致しません: %s", e)
        raise

    sheet = wb.sheets[0]

    quotation_no = sheet.range("S1").value
    quotation_project = sheet.range("C6").value
    quotation_customer = sheet.range("C1").value

    # ASC変換
    if "ASC" in [s.name for s in wb.sheets]:
        wb.sheets["ASC"].delete()
    asc_sheet = wb.sheets.add(name="ASC", after=wb.sheets[-1])
    asc_sheet.range("A1").value = quotation_customer
    asc_sheet.range("A2").formula = "=ASC(A1)"
    customer_text = asc_sheet.range("A2").value

    wb.sheets[0].activate()

    logging.info("✅ Excel情報の読み取り完了")
    return {
        "quotation_no": quotation_no,
        "quotation_project": quotation_project,
        "quotation_customer": quotation_customer,
        "customer_text": customer_text
    }

# ✅ 動作確認用（直接実行時のみ有効）
if __name__ == "__main__":
    import sys
    from pathlib import Path

    # 📁 config.py のある親フォルダをパスに追加
    sys.path.append(str(Path(__file__).resolve().parent.parent))

    # 🔁 再インポート（__main__ 限定）
    from config import EXCEL_FILE_NAME

    # ログ出力設定（簡易版）
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    from pprint import pprint
    result = extract_excel_info()

    # 絵文字付きラベルに変換
    display_data = {
        "📄 quotation_No": result["quotation_no"],
        "📝 project_name": result["quotation_project"],
        "🏢 customer": result["quotation_customer"],
        "🔤 customer_text": result["customer_text"]
    }

    print(tabulate(display_data.items(), headers=["項目", "値"], tablefmt="grid"))