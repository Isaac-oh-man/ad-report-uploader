"""Google Sheets 업로드 모듈."""
import math
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from config import (
    GOOGLE_SHEETS_ID,
    GOOGLE_SERVICE_ACCOUNT_FILE,
    WORKSHEET_NAME,
    SHEET_COLUMNS,
    DATA_START_ROW,
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 숫자로 올려야 하는 컬럼
NUMERIC_COLS = {"impression", "clicks", "cost", "view"}


def _get_client() -> gspread.Client:
    creds = Credentials.from_service_account_file(
        GOOGLE_SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    return gspread.authorize(creds)


def _get_worksheet() -> gspread.Worksheet:
    client = _get_client()
    spreadsheet = client.open_by_key(GOOGLE_SHEETS_ID)
    return spreadsheet.worksheet(WORKSHEET_NAME)


def _find_next_empty_row(ws: gspread.Worksheet) -> int:
    col_b = ws.col_values(2)
    return len(col_b) + 1


def _to_native(val):
    """NaN/None -> '', 숫자는 int/float 유지, 나머지는 str."""
    if val is None:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    if isinstance(val, (int, float)):
        return int(val)
    return str(val)


def _build_rows(df: pd.DataFrame) -> list[list]:
    """DataFrame을 시트에 올릴 2D 리스트로 변환. 숫자는 숫자 타입 유지."""
    rows = []
    for _, row in df.iterrows():
        r = []
        for col in SHEET_COLUMNS:
            val = row.get(col, "")
            r.append(_to_native(val))
        rows.append(r)
    return rows


def upload_to_sheet(df: pd.DataFrame, brand: str = "default", mode: str = "append"):
    if df.empty:
        print("[WARN] 업로드할 데이터가 없습니다.")
        return

    ws = _get_worksheet()
    rows = _build_rows(df)

    if mode == "overwrite":
        end_col = chr(ord("B") + len(SHEET_COLUMNS) - 1)
        clear_range = f"B{DATA_START_ROW}:{end_col}5000"
        ws.batch_clear([clear_range])
        cell = f"B{DATA_START_ROW}"
        ws.update(range_name=cell, values=rows, value_input_option="RAW")
        print(f"[OK] {len(rows)}행 덮어쓰기 완료 (B{DATA_START_ROW}부터)")
    else:
        next_row = _find_next_empty_row(ws)
        cell = f"B{next_row}"
        ws.update(range_name=cell, values=rows, value_input_option="RAW")
        print(f"[OK] {len(rows)}행 추가 완료 (B{next_row}부터)")
