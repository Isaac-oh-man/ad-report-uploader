"""광고 보고서 -> 구글 시트 업로드 (Streamlit Web UI)."""
import json, os, re, math, io
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BRANDS_FILE = os.path.join(BASE_DIR, "brands.json")
CRED_FILE = os.path.join(BASE_DIR, "credentials", "service_account.json")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
           "https://www.googleapis.com/auth/drive"]
CHANNELS = ["GoogleAds", "MetaBusiness", "KakaoMoment", "NaverGFA"]
PLATFORM_ORDER = {c: i for i, c in enumerate(CHANNELS)}
STD_COLS = ["reportDate","campaignName","assetName","adPlatformType",
            "impression","clicks","cost","view"]

# ── 유틸 ──────────────────────────────────────────
def _load_brands() -> dict:
    if os.path.exists(BRANDS_FILE):
        with open(BRANDS_FILE, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    return {}

def _save_brands(d: dict):
    with open(BRANDS_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def _sheet_id_from_link(link: str) -> str:
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", link)
    return m.group(1) if m else link.strip()

def _gs_client() -> gspread.Client:
    if "gcp_service_account" in st.secrets:
        info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(CRED_FILE, scopes=SCOPES)
    return gspread.authorize(creds)

def _to_int(v) -> int:
    try: return int(float(str(v).replace(",","")))
    except: return 0

def _to_native(v):
    if v is None: return ""
    if isinstance(v, float) and math.isnan(v): return ""
    if isinstance(v, (int, float)): return int(v)
    return str(v)

def _parse_gfa_date(s: str) -> str:
    c = s.strip().rstrip(".")
    p = c.split(".")
    return f"{p[0]}-{p[1]}-{p[2]}" if len(p)==3 else c


# ── CSV 인코딩 자동감지 ───────────────────────────
def _read_csv_auto(fbytes) -> pd.DataFrame:
    """UTF-8, UTF-8-SIG, CP949 순서로 시도하여 CSV를 읽는다."""
    raw = fbytes.read()
    fbytes.seek(0)
    for enc in ["utf-8-sig", "utf-8", "cp949", "euc-kr"]:
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc)
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    return pd.read_csv(io.BytesIO(raw), encoding="utf-8", errors="replace")


# ── 시트 헤더 / 중복 체크 ─────────────────────────
def _fetch_sheet_headers(sheet_id, ws_name, header_row=11):
    ws = _gs_client().open_by_key(sheet_id).worksheet(ws_name)
    row = ws.row_values(header_row)
    return [c.strip() for c in row if c.strip()]

def _get_existing_dates(cfg: dict) -> set:
    """시트에서 이미 업로드된 날짜 목록을 가져온다."""
    col_map = cfg.get("column_map", {})
    date_letter = None
    for letter, col_name in col_map.items():
        if col_name == "reportDate":
            date_letter = letter
            break
    if not date_letter:
        return set()
    col_idx = ord(date_letter) - ord("A") + 1
    ws = _gs_client().open_by_key(cfg["sheet_id"]).worksheet(cfg["worksheet"])
    vals = ws.col_values(col_idx)
    return set(v.strip() for v in vals if v.strip() and v.strip() != "reportDate")


# ── 시트 업로드 (컬럼 매핑) ───────────────────────
def _upload_mapped(df, cfg):
    col_map = cfg.get("column_map", {})
    if not col_map:
        st.error("컬럼 매핑이 설정되지 않았습니다.")
        return 0, 0
    col_letters = sorted(col_map.keys())
    client = _gs_client()
    ws = client.open_by_key(cfg["sheet_id"]).worksheet(cfg["worksheet"])
    first_col_idx = ord(col_letters[0]) - ord("A") + 1
    existing = ws.col_values(first_col_idx)
    next_row = len(existing) + 1
    first_letter = col_letters[0]
    last_letter = col_letters[-1]
    num_cols = ord(last_letter) - ord(first_letter) + 1
    rows = []
    for _, row in df.iterrows():
        r = [""] * num_cols
        for letter, std_col in col_map.items():
            idx = ord(letter) - ord(first_letter)
            r[idx] = _to_native(row.get(std_col, ""))
        rows.append(r)
    cell = f"{first_letter}{next_row}"
    ws.update(range_name=cell, values=rows, value_input_option="RAW")
    return next_row, len(rows)


# ── 파일 로더 ─────────────────────────────────────
def _load_gfa(df_raw, cfg):
    df = df_raw.copy()
    df.columns = df.columns.str.strip()
    asset_map = cfg.get("asset_map", {})
    divisor = cfg.get("cost_divisor", 0.605)
    if asset_map:
        def match(name):
            for key, val in asset_map.items():
                if all(p.strip() in name for p in key.split("|")):
                    return val
            return None
        df["assetName"] = df["광고 소재 이름"].apply(match)
        df = df.dropna(subset=["assetName"])
    else:
        df["assetName"] = df["광고 소재 이름"]
    df["reportDate"] = df["기간"].apply(_parse_gfa_date)
    df["cost_calc"] = pd.to_numeric(df["총비용"], errors="coerce").fillna(0) / divisor
    df["impression"] = pd.to_numeric(df["노출수"], errors="coerce").fillna(0)
    df["clicks"] = pd.to_numeric(df["클릭수"], errors="coerce").fillna(0)
    df["view"] = pd.to_numeric(df["총 재생수"], errors="coerce").fillna(0)
    g = df.groupby(["reportDate","assetName"]).agg(
        impression=("impression","sum"), clicks=("clicks","sum"),
        cost=("cost_calc","sum"), view=("view","sum")).reset_index()
    g["campaignName"] = cfg.get("campaign_name","")
    g["adPlatformType"] = "NaverGFA"
    for c in ["impression","clicks","cost","view"]:
        g[c] = g[c].apply(_to_int)
    return g[STD_COLS]

def _load_excel(df_raw):
    df = df_raw.copy()
    df.columns = df.columns.str.strip()
    col_map = {"Date":"reportDate","CampaignName":"campaignName",
               "AssetName":"assetName","AdPlatformType":"adPlatformType",
               "Impression":"impression","Click":"clicks",
               "Cost":"cost","View":"view"}
    df = df.rename(columns=col_map)
    for c in ["impression","clicks","cost","view"]:
        if c in df.columns: df[c] = df[c].apply(_to_int)
    return df[STD_COLS]

def _detect_load(fname, fbytes, cfg):
    ext = os.path.splitext(fname)[1].lower()
    if ext == ".csv":
        return _load_gfa(_read_csv_auto(fbytes), cfg)
    elif ext in (".xlsx",".xls"):
        return _load_excel(pd.read_excel(fbytes))
    return pd.DataFrame()


# ══════════════════════════════════════════════════
# Streamlit UI
# ══════════════════════════════════════════════════
st.set_page_config(page_title="광고 리포트 업로더", page_icon="📊", layout="wide")
st.title("📊 광고 리포트 → 구글 시트 업로더")

brands = _load_brands()
MAPPING_OPTIONS = ["(매핑 안함)"] + STD_COLS
tab_guide, tab_upload, tab_settings = st.tabs(["📖 가이드", "📤 업로드", "⚙️ 브랜드 설정"])

# ── 가이드 탭 ─────────────────────────────────────
with tab_guide:
    st.subheader("사용 가이드")
    st.markdown("""
이 앱은 **광고 매체별 성과 보고서를 구글 스프레드시트에 자동 업로드**하는 도구입니다.

---

### 📌 사용 순서

1. **브랜드 설정** (최초 1회)
   - "브랜드 설정" 탭에서 브랜드명, 스프레드시트 링크, 워크시트 이름 등을 입력
   - "시트 헤더 불러오기"로 컬럼 매핑 설정
   - 네이버 GFA 사용 시 마진 계수 및 소재 매핑 설정

2. **파일 업로드**
   - "업로드" 탭에서 브랜드 선택 후 파일 업로드
   - 미리보기에서 데이터 확인
   - "구글 시트에 업로드" 버튼 클릭

---

### 📁 파일 형식 안내

| 채널 | 파일 출처 | 형식 |
|------|----------|------|
| **네이버 GFA** | 네이버 광고 매체 대시보드에서 성과 보고서 다운로드 | CSV (ZIP 해제 후) |
| **구글 애즈 / 메타 / 카카오모먼트** | TW360 콘솔 보고서 파일 활용 | Excel 또는 CSV 모두 가능 |

> 💡 TW360에서 CSV로 다운로드 시 한글이 깨지는 경우가 있지만, 이 앱에서 자동으로 인코딩을 감지하여 처리합니다.

---

### ⚠️ 주의사항

- **중복 업로드 방지**: 이미 시트에 존재하는 날짜의 데이터를 업로드하려고 하면 경고가 표시됩니다.
- **컬럼 매핑**: 브랜드마다 스프레드시트 구조가 다를 수 있으므로, 반드시 "시트 헤더 불러오기"로 매핑을 설정하세요.
- **네이버 GFA 비용**: 마진 계수(기본 0.605 = 0.55 x 1.1)가 자동 적용됩니다.
- **정렬 순서**: 업로드 시 날짜 → GoogleAds → MetaBusiness → KakaoMoment → NaverGFA 순으로 정렬됩니다.
""")


# ── 업로드 탭 (중복 체크 포함) ────────────────────
with tab_upload:
    if not brands:
        st.warning("먼저 '브랜드 설정' 탭에서 브랜드를 등록하세요.")
    else:
        sel_brand = st.selectbox("브랜드 선택", list(brands.keys()))
        cfg = brands[sel_brand]
        uploaded = st.file_uploader(
            "파일 업로드 (CSV / Excel, 여러 파일 가능)",
            type=["csv","xlsx","xls"], accept_multiple_files=True)

        if uploaded:
            all_dfs = []
            for f in uploaded:
                df = _detect_load(f.name, f, cfg)
                if not df.empty:
                    all_dfs.append(df)

            if all_dfs:
                combined = pd.concat(all_dfs, ignore_index=True)
                combined["_ord"] = combined["adPlatformType"].map(PLATFORM_ORDER).fillna(9)
                combined = combined.sort_values(
                    ["reportDate","_ord","assetName"]
                ).drop(columns=["_ord"]).reset_index(drop=True)

                # 중복 날짜 체크
                dup_warning = False
                try:
                    existing_dates = _get_existing_dates(cfg)
                    new_dates = set(combined["reportDate"].unique())
                    overlap = new_dates & existing_dates
                    if overlap:
                        dup_warning = True
                        st.error(f"⚠️ 다음 날짜는 이미 시트에 존재합니다: **{', '.join(sorted(overlap))}**\n\n중복 업로드를 방지하기 위해 해당 날짜를 확인해주세요.")
                except Exception:
                    pass  # 체크 실패 시 무시하고 진행 허용

                st.subheader("📋 미리보기")
                dates = combined["reportDate"]
                st.markdown(f"**{len(combined)}행** | {dates.min()} ~ {dates.max()}")
                channels = combined["adPlatformType"].value_counts()
                metric_cols = st.columns(len(channels))
                for i, (ch, cnt) in enumerate(channels.items()):
                    metric_cols[i].metric(ch, f"{cnt}행")
                st.dataframe(combined, use_container_width=True, height=400)

                st.markdown("---")
                if dup_warning:
                    st.warning("중복 날짜가 감지되었습니다. 그래도 업로드하려면 아래 체크박스를 선택하세요.")
                    force = st.checkbox("중복 날짜 무시하고 업로드")
                else:
                    force = True

                if force and st.button("🚀 구글 시트에 업로드", type="primary"):
                    with st.spinner("업로드 중..."):
                        try:
                            start, cnt = _upload_mapped(combined, cfg)
                            st.success(f"✅ {cnt}행 업로드 완료! (행 {start}부터)")
                            link = f"https://docs.google.com/spreadsheets/d/{cfg['sheet_id']}"
                            st.markdown(f"[📎 스프레드시트 열기]({link})")
                        except Exception as e:
                            st.error(f"❌ 업로드 실패: {e}")
            else:
                st.warning("처리 가능한 데이터가 없습니다.")


# ── 브랜드 설정 탭 ────────────────────────────────
with tab_settings:
    st.subheader("브랜드 추가 / 수정")
    brand_name = st.text_input("브랜드명", placeholder="예: 노랑통닭")
    sheet_link = st.text_input("구글 스프레드시트 링크")
    ws_name = st.text_input("워크시트 이름", placeholder="예: 26.04")
    header_row = st.number_input("헤더 행 번호", value=11, min_value=1, step=1)
    campaign_name = st.text_input("campaignName", placeholder="예: TW360.노랑통닭_CPV")

    st.markdown("---")
    st.markdown("**📋 컬럼 매핑** — 시트 헤더를 불러와서 각 컬럼의 역할을 지정하세요")
    if st.button("🔄 시트 헤더 불러오기"):
        if sheet_link and ws_name:
            try:
                sid = _sheet_id_from_link(sheet_link)
                hdrs = _fetch_sheet_headers(sid, ws_name, int(header_row))
                st.session_state["fetched_headers"] = hdrs
                st.success(f"헤더 {len(hdrs)}개 불러옴: {hdrs}")
            except Exception as e:
                st.error(f"헤더 불러오기 실패: {type(e).__name__}: {e}")
        else:
            st.warning("스프레드시트 링크와 워크시트 이름을 먼저 입력하세요.")

    column_map = {}
    if "fetched_headers" in st.session_state:
        hdrs = st.session_state["fetched_headers"]
        st.markdown("각 시트 컬럼에 대응하는 데이터를 선택하세요:")
        cols_per_row = 4
        for i in range(0, len(hdrs), cols_per_row):
            row_cols = st.columns(cols_per_row)
            for j, col_st in enumerate(row_cols):
                idx = i + j
                if idx < len(hdrs):
                    h = hdrs[idx]
                    letter = chr(ord("A") + idx + 1)
                    default_idx = 0
                    for k, opt in enumerate(MAPPING_OPTIONS):
                        if opt.lower() == h.lower():
                            default_idx = k
                            break
                    sel = col_st.selectbox(f"{letter}열: {h}", MAPPING_OPTIONS, index=default_idx, key=f"map_{idx}")
                    if sel != "(매핑 안함)":
                        column_map[letter] = sel

    st.markdown("---")
    st.markdown("**네이버 GFA 설정** (해당 시에만)")
    use_naver = st.checkbox("네이버 GFA 사용")
    cost_divisor = st.number_input("마진 계수 (총비용 / 이 값)", value=0.605, step=0.001, format="%.3f")
    is_video = st.checkbox("동영상 캠페인 (소재 매핑 필요)")
    asset_map_text = st.text_area("소재 매핑 (키워드1|키워드2 = 시트 소재명, 한 줄씩)",
        placeholder="키즈팝편|15s = 우도 땅콩 치킨_키즈팝편_15s", height=100)

    if st.button("💾 브랜드 저장", type="primary"):
        if not brand_name:
            st.warning("브랜드명을 입력하세요.")
        else:
            asset_map = {}
            if is_video and asset_map_text.strip():
                for line in asset_map_text.strip().split("\n"):
                    if "=" in line:
                        k, v = line.split("=", 1)
                        asset_map[k.strip()] = v.strip()
            brands[brand_name] = {
                "sheet_id": _sheet_id_from_link(sheet_link) if sheet_link else "",
                "worksheet": ws_name, "header_row": int(header_row),
                "campaign_name": campaign_name, "column_map": column_map,
                "use_naver": use_naver, "cost_divisor": cost_divisor,
                "is_video": is_video, "asset_map": asset_map,
            }
            _save_brands(brands)
            st.success(f"'{brand_name}' 저장 완료!")
            st.rerun()

    if brands:
        st.markdown("---")
        st.subheader("등록된 브랜드")
        for name, cfg in brands.items():
            with st.expander(f"📁 {name}"):
                st.json(cfg)
                if st.button(f"🗑️ {name} 삭제", key=f"del_{name}"):
                    del brands[name]
                    _save_brands(brands)
                    st.rerun()
