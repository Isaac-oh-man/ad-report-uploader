import os
from dotenv import load_dotenv

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(_BASE_DIR, ".env"))

# --- CSV 감시 폴더 ---
WATCH_DIR = os.getenv(
    "WATCH_DIR",
    os.path.join(_BASE_DIR, "inbox"),
)

# --- 처리 완료 파일 이동 폴더 ---
DONE_DIR = os.getenv(
    "DONE_DIR",
    os.path.join(_BASE_DIR, "done"),
)

# --- Google Sheets ---
GOOGLE_SHEETS_ID = os.getenv(
    "GOOGLE_SHEETS_ID",
    "1MHMHKbdPKtf6RijQdWjPzTxA-z72XHDd_jxJ6xyUGdk",
)

_default_cred = os.path.join(_BASE_DIR, "credentials", "service_account.json")
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv(
    "GOOGLE_SERVICE_ACCOUNT_FILE", _default_cred
)

# --- 대상 워크시트명 ---
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME", "26.04")

# --- 업로드 컬럼 (B~I열, view까지만. CPV~CTR은 시트 수식) ---
SHEET_COLUMNS = [
    "reportDate",
    "campaignName",
    "assetName",
    "adPlatformType",
    "impression",
    "clicks",
    "cost",
    "view",
]

# --- 데이터 시작 행 ---
DATA_START_ROW = 13
