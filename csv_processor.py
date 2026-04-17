"""파일 처리 모듈: GFA CSV + 구글/메타 Excel -> 시트 형식 변환."""
import os
import shutil
import zipfile
import pandas as pd
from config import SHEET_COLUMNS, WATCH_DIR, DONE_DIR

# --- 네이버 GFA 소재 키워드 -> assetName 매핑 ---
ASSET_MAP = {
    ("키즈팝편", "15s"): "우도 땅콩 치킨_키즈팝편_15s",
    ("키즈팝편", "30s"): "우도 땅콩 치킨_키즈팝편_30s",
    ("힙합편", "15s"): "우도땅콩치킨 힙합편_15s",
    ("힙합편", "30s"): "우도땅콩치킨 힙합편_30s",
}

CAMPAIGN_NAME = "TW360.노랑통닭_CPV"
GFA_COST_DIVISOR = 0.55 * 1.1


def ensure_dirs():
    os.makedirs(WATCH_DIR, exist_ok=True)
    os.makedirs(DONE_DIR, exist_ok=True)


def _to_int(val) -> int:
    try:
        return int(float(str(val).replace(",", "")))
    except (ValueError, TypeError):
        return 0


def _match_asset(name: str) -> str | None:
    for (keyword, duration), asset_name in ASSET_MAP.items():
        if keyword in name and duration in name:
            return asset_name
    return None


def _parse_gfa_date(date_str: str) -> str:
    clean = date_str.strip().rstrip(".")
    parts = clean.split(".")
    if len(parts) == 3:
        return f"{parts[0]}-{parts[1]}-{parts[2]}"
    return clean


def _load_gfa(csv_path: str) -> pd.DataFrame:
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    df.columns = df.columns.str.strip()

    df["assetName"] = df["광고 소재 이름"].apply(_match_asset)
    df = df.dropna(subset=["assetName"])
    df["reportDate"] = df["기간"].apply(_parse_gfa_date)

    df["cost_raw"] = pd.to_numeric(df["총비용"], errors="coerce").fillna(0)
    df["cost_calc"] = df["cost_raw"] / GFA_COST_DIVISOR
    df["impression"] = pd.to_numeric(df["노출수"], errors="coerce").fillna(0)
    df["clicks"] = pd.to_numeric(df["클릭수"], errors="coerce").fillna(0)
    df["view"] = pd.to_numeric(df["총 재생수"], errors="coerce").fillna(0)

    grouped = df.groupby(["reportDate", "assetName"]).agg(
        impression=("impression", "sum"),
        clicks=("clicks", "sum"),
        cost=("cost_calc", "sum"),
        view=("view", "sum"),
    ).reset_index()

    grouped["campaignName"] = CAMPAIGN_NAME
    grouped["adPlatformType"] = "NaverGFA"
    return grouped


def _load_report_excel(xlsx_path: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path)
    df.columns = df.columns.str.strip()

    col_map = {
        "Date": "reportDate",
        "CampaignName": "campaignName",
        "AssetName": "assetName",
        "AdPlatformType": "adPlatformType",
        "Impression": "impression",
        "Click": "clicks",
        "Cost": "cost",
        "View": "view",
    }
    df = df.rename(columns=col_map)

    for c in ["impression", "clicks", "cost", "view"]:
        df[c] = pd.to_numeric(
            df[c].astype(str).str.replace(",", ""), errors="coerce"
        ).fillna(0)

    return df[["reportDate", "campaignName", "assetName",
               "adPlatformType", "impression", "clicks", "cost", "view"]]


def load_file(file_path: str) -> pd.DataFrame:
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".csv":
        raw = _load_gfa(file_path)
    elif ext in (".xlsx", ".xls"):
        raw = _load_report_excel(file_path)
    else:
        print(f"[SKIP] 지원하지 않는 형식: {file_path}")
        return pd.DataFrame()

    if raw.empty:
        return raw

    # 숫자를 정수로 변환 (포맷 없이 순수 숫자)
    for c in ["impression", "clicks", "cost", "view"]:
        raw[c] = raw[c].apply(_to_int)

    return raw[SHEET_COLUMNS]


def extract_files(file_path: str) -> list[str]:
    if file_path.endswith(".zip"):
        with zipfile.ZipFile(file_path, "r") as zf:
            names = [f for f in zf.namelist()
                     if f.endswith(".csv") or f.endswith(".xlsx") or f.endswith(".xls")]
            zf.extractall(os.path.dirname(file_path))
        return [os.path.join(os.path.dirname(file_path), n) for n in names]
    else:
        return [file_path]


def move_to_done(file_path: str):
    dest = os.path.join(DONE_DIR, os.path.basename(file_path))
    shutil.move(file_path, dest)
    print(f"[OK] {os.path.basename(file_path)} -> done/")
