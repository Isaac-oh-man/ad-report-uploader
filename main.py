"""GFA + 구글/메타 보고서 -> 구글 시트 자동 업로드."""
import argparse
import glob
import os
import pandas as pd
from datetime import datetime

from csv_processor import ensure_dirs, extract_files, load_file, move_to_done
from sheet_uploader import upload_to_sheet
from config import WATCH_DIR, SHEET_COLUMNS

PLATFORM_ORDER = {"GoogleAds": 0, "MetaBusiness": 1, "NaverGFA": 2}


def process_inbox(mode: str):
    patterns = ["*.zip", "*.csv", "*.xlsx", "*.xls"]
    files = []
    for p in patterns:
        files += glob.glob(os.path.join(WATCH_DIR, p))
    if not files:
        print("[INFO] inbox에 처리할 파일이 없습니다.")
        return

    # 모든 파일 로드 후 합치기
    all_dfs = []
    processed = []
    for f in sorted(files):
        print(f"\n[LOAD] {os.path.basename(f)}")
        inner_files = extract_files(f)
        for inner in inner_files:
            df = load_file(inner)
            if not df.empty:
                print(f"  -> {len(df)}행 로드")
                all_dfs.append(df)
            if inner != f:
                processed.append(inner)
        processed.append(f)

    if not all_dfs:
        print("[INFO] 업로드할 데이터가 없습니다.")
        return

    # 합산 후 정렬: 날짜 -> 플랫폼(G/M/N) -> 소재명
    combined = pd.concat(all_dfs, ignore_index=True)
    combined["_platform_order"] = combined["adPlatformType"].map(PLATFORM_ORDER).fillna(9)
    combined = combined.sort_values(
        ["reportDate", "_platform_order", "assetName"]
    ).reset_index(drop=True)
    combined = combined[SHEET_COLUMNS]

    # 한번에 업로드
    print(f"\n[UPLOAD] 총 {len(combined)}행")
    upload_to_sheet(combined, mode=mode)

    # 처리 완료 파일 이동
    for f in processed:
        move_to_done(f)


def main():
    parser = argparse.ArgumentParser(description="광고 보고서 -> 구글 시트 업로드")
    parser.add_argument("--mode", default="append", choices=["append", "overwrite"])
    args = parser.parse_args()

    ensure_dirs()

    print(f"{'='*50}")
    print(f"  광고 보고서 -> 구글 시트 업로드")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  inbox: {WATCH_DIR}")
    print(f"{'='*50}")

    process_inbox(args.mode)
    print("\n완료!")


if __name__ == "__main__":
    main()
