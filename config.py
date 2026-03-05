"""
Streamlitアプリの設定ファイル
認証情報は入力フォームから取得するため、.envには不要
"""

import os
from pathlib import Path

# ============ SharePoint設定 ============
SHAREPOINT_SITE_URL = "https://toyodagoseicorp.sharepoint.com/sites/FC644"
SHAREPOINT_LIBRARY = "Shared Documents"
EXCEL_FILENAME = "メンバー管理表.xlsx"

# ============ データ処理設定 ============
INCLUDE_KEYWORDS = ["ＦＣ技", "出向"]
EXCLUDE_KEYWORDS = ["ＴＧテクノ"]
MAIL_FILTER_DOMAIN = "@ts"
派遣_ROLE = "派遣・請負社員"

# ============ ファイルパス設定 ============
DATA_DIR = "data"
INPUT_DIR = os.path.join(DATA_DIR, "input")
OUTPUT_DIR = os.path.join(DATA_DIR, "output")
TEMP_DIR = os.path.join(DATA_DIR, "temp")
OUTPUT_FILENAME = "名簿リスト.csv"
OUTPUT_ENCODING = "utf-8-sig"

# ディレクトリ作成
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)

# ============ 列名マッピング ============
COLUMN_MAPPING = {
    "department": "部署",
    "group": "Gr",
    "team": "Tm",
    "position": "役職",
    "name": "名前",
    "employee_code": "社員コード"
}

# ============ ソース列名 ============
SOURCE_COLUMNS = {
    "department": "部署",
    "mail": "メール",
    "position": "役職",
    "last_name": "姓",
    "first_name": "名",
    "nickname": "ニックネーム"
}

# ============ Streamlit設定 ============
APP_TITLE = "メンバー管理表 → 名簿リスト 変換ツール"
APP_ICON = "👥"
PAGE_LAYOUT = "wide"
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB

# ============ 認証設定 ============
AUTH_TIMEOUT = 30  # 秒
SESSION_TIMEOUT = 3600  # 1時間