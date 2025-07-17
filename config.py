# 📌 config.py - 共通設定管理ファイル

from pathlib import Path
from datetime import datetime
import logging
from pathlib import Path



# 📁 ログ出力先フォルダ（存在しない場合は作成）
logs_dir = Path("./logs")
logs_dir.mkdir(exist_ok=True)

# 🪵 ログ初期設定（コンソール ＋ ファイル）
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),  # コンソール出力
        logging.FileHandler(logs_dir / "app.log", encoding="utf-8")  # ファイル出力
    ]
)

# 📜 ログメッセージ出力
logging.info("🚀 処理を開始しました")

# 💾 提出用フォルダの親ディレクトリ
QUOTATION_BASE_DIR = Path(
    r"C:\Users\2012003\OneDrive - 株式会社川本製作所\2012003_福田\01_営業\10_見積もり\01_提出見積もり")


# 🎯 対象のExcelファイル名（既に開いている前提）
EXCEL_FILE_NAME = "ASKA101P11.xlsm"

# 📅 フォルダ名に使う日付フォーマット
DATE_FORMAT = "%Y%m%d"


# 🏷 フォルダ名のテンプレート
# 例）20250709【Q】KMT-7619-017318-01_NHAT_Sewage pump cable for ZUJ and float switches EHF-5

# 📏 案件名（プロジェクト名）の最大文字数
MAX_PROJECT_NAME_LENGTH = 50

def generate_folder_name(
    quotation_no: str,
    customer_text: str,
    project_name: str,
    max_project_name_length: int = MAX_PROJECT_NAME_LENGTH
) -> str:
    
    date_str = datetime.now().strftime(DATE_FORMAT)
    import re

    # 禁止文字の置換
    safe_project_name = re.sub(r'[\\/:*?"<>|]', '_', project_name)

    # ⛏ 長すぎる場合は切り捨て（末尾カット）
    if len(safe_project_name) > max_project_name_length:
        safe_project_name = safe_project_name[:max_project_name_length].rstrip()

    folder_name = f"{date_str}【Q】{quotation_no}_{customer_text}_{safe_project_name}"

    
    # 🪵 ログ出力（フォルダ名の生成確認）
    logging.info("フォルダ名を生成しました: %s", folder_name)
    
    return folder_name

# ✅ 動作確認用テストブロック
if __name__ == "__main__":
    quotation_no = "KMT-7619-017318-01"
    customer_text = "NHAT"
    project_name = "Sewage pump cable for ZUJ and float switches EHF-5"

    result = generate_folder_name(quotation_no, customer_text, project_name)
    print("📂 生成フォルダ名：", result)