import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, date, timedelta
import requests
from bs4 import BeautifulSoup
import re
import time
import io
import os
import json
import subprocess
import zipfile
import FinanceDataReader as fdr
from PIL import Image


# ── Playwright Chromium 자동 설치 (Streamlit Cloud) ──────────
@st.cache_resource
def ensure_playwright_browser():
    """Streamlit Cloud에서 Chromium 자동 설치. 결과 메시지 반환."""
    msgs = []
    try:
        result = subprocess.run(
            ["playwright", "install", "chromium"],
            capture_output=True, timeout=300, text=True,
        )
        msgs.append(f"install rc={result.returncode}")
        if result.stderr:
            msgs.append(f"stderr: {result.stderr[-300:]}")
    except Exception as e:
        msgs.append(f"install exception: {e}")

    # 실제로 동작하는지 테스트
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
            b.close()
        msgs.append("launch OK")
        return True, msgs
    except Exception as e:
        msgs.append(f"launch failed: {e}")
        return False, msgs

st.set_page_config(page_title="종가 캡처", layout="centered")
st.title("코스닥/코스피 종가 조회")

# ── 종목코드 조회 ─────────────────────────────────────────────
def clean_stock_name(name: str) -> str:
    return re.sub(r'[㈜㈔\s]', '', name).strip()


# 투자명부 표기와 실제 상장명이 달라 자동 매칭이 안 되는 종목들 (발견될 때마다 추가)
STOCK_NAME_ALIASES = {
    "뉴엔에이아이": "뉴엔AI",
}


@st.cache_data(ttl=3600 * 12, show_spinner="전종목 코드 로딩 중 (최초 1회)...")
def get_code_map() -> dict:
    """FinanceDataReader로 코스피+코스닥 전종목 코드맵 반환."""
    code_map = {}
    for market in ["KOSPI", "KOSDAQ"]:
        try:
            df = fdr.StockListing(market)
            for _, row in df.iterrows():
                name = str(row.get("Name", "")).strip()
                code = str(row.get("Code", "")).strip()
                if name and code:
                    code_map[name] = code
                    code_map[clean_stock_name(name)] = code
        except Exception:
            pass
    return code_map


def get_stock_code(name: str, code_map: dict) -> str | None:
    name = name.strip()
    code = code_map.get(name) or code_map.get(clean_stock_name(name))
    if code:
        return code
    alias = STOCK_NAME_ALIASES.get(clean_stock_name(name))
    if alias:
        return code_map.get(alias) or code_map.get(clean_stock_name(alias))
    return None


# ── 투자명부 자동 리스트업 ───────────────────────────────────
LEDGER_DIR = r"G:\공유 드라이브\4-2. 기획본부_투자관리팀\01. 투자명부"
FUNDS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "funds")

# 투자명부 '계정명' ↔ 관리 펀드(funds/*.json) 매핑. 계정명이 없으면 None.
FUND_ACCOUNT_MAP = {
    "2013_성장사다리": None,
    "2017_4차산업": "4차 산업혁명",
    "2020_소부장": "소부장",
    "2021_이노베이션": "이노베이션",
    "2024_청년창업": "청년창업",
    "2024_기술혁신전문": "기술혁신",
}


def load_fund_names() -> dict:
    """funds/*.json에서 fund_id -> 표시용 fund_name 매핑."""
    names = {}
    if os.path.isdir(FUNDS_DIR):
        for fname in os.listdir(FUNDS_DIR):
            if not fname.endswith(".json"):
                continue
            fid = fname[:-5]
            try:
                with open(os.path.join(FUNDS_DIR, fname), encoding="utf-8") as f:
                    data = json.load(f)
                names[fid] = data.get("fund_name", fid)
            except Exception:
                names[fid] = fid
    return names


def find_ledger_file(target_date: date, max_back_months: int = 12) -> tuple[str | None, str | None]:
    """target_date가 속한 달부터 역순으로 투자명부(YYYY년MM월말기준).xlsx 탐색."""
    if not os.path.isdir(LEDGER_DIR):
        return None, None
    y, m = target_date.year, target_date.month
    for _ in range(max_back_months):
        fname = f"투자명부({y}년{m:02d}월말기준).xlsx"
        fpath = os.path.join(LEDGER_DIR, fname)
        if os.path.isfile(fpath):
            return fpath, f"{y}년{m:02d}월말기준"
        m -= 1
        if m == 0:
            m, y = 12, y - 1
    return None, None


def extract_fund_stock_list(ledger_path: str, code_map: dict) -> tuple[pd.DataFrame, list[str]]:
    """투자명부에서 잔고가 있는 모든 계정(펀드)의 상장종목(AS열 상장여부='Y')을 추출한다.
    관리 펀드(FUND_ACCOUNT_MAP)로 매핑되는 계정은 등록된 펀드명으로 표시하고,
    그 외 계정(다른 조합/본계정 등)도 계정명 그대로 포함한다.
    """
    wb = openpyxl.load_workbook(ledger_path, data_only=True)
    ws = wb["투자명부"]
    headers = [ws.cell(row=4, column=c).value for c in range(1, ws.max_column + 1)]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}
    col_account = col_idx.get("계정명")
    col_type = col_idx.get("구분")
    col_company = col_idx.get("투자업체명")
    col_listed = col_idx.get("상장여부")

    if not (col_account and col_type and col_company and col_listed):
        return pd.DataFrame(), ["투자명부 시트에서 계정명/구분/투자업체명/상장여부 열을 찾지 못했습니다."]

    fund_names = load_fund_names()
    account_to_fund_id = {account: fund_id for fund_id, account in FUND_ACCOUNT_MAP.items() if account}

    by_account: dict[str, set] = {}
    for r in range(5, ws.max_row + 1):
        account = ws.cell(row=r, column=col_account).value
        if not account:
            continue
        if (ws.cell(row=r, column=col_type).value == "잔고"
                and str(ws.cell(row=r, column=col_listed).value).strip().upper() == "Y"):
            name = ws.cell(row=r, column=col_company).value
            if name:
                by_account.setdefault(account, set()).add(str(name).strip())

    notices = []
    rows = []
    for account, companies in sorted(by_account.items()):
        fund_id = account_to_fund_id.get(account)
        display_name = fund_names.get(fund_id, fund_id) if fund_id else account

        no_code = [c for c in companies if not get_stock_code(c, code_map)]
        if no_code:
            notices.append(f"{display_name}: 상장종목이나 종목코드 미매칭 — {', '.join(sorted(no_code))} (수동 확인 필요)")
        for name in sorted(companies):
            rows.append({"fund": display_name, "name": name})

    return pd.DataFrame(rows), notices


# ── 종가 조회 ────────────────────────────────────────────────
# 네이버가 finance.naver.com/item/* 페이지를 stock.naver.com으로 전면 이전하면서
# 기존 sise_day 스크래핑이 폐지됨(HTTP 410). 모바일 API로 대체.
PRICE_PAGE_SIZE = 60


def _fetch_price_page(code: str, page: int) -> list[dict] | None:
    url = f"https://m.stock.naver.com/api/stock/{code}/price?page={page}&pageSize={PRICE_PAGE_SIZE}"
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        if res.status_code != 200:
            return None
        rows = res.json()
        return rows or None
    except Exception:
        return None


def fetch_closing_price(code: str, target_date: date) -> int | None:
    date_str = target_date.strftime("%Y-%m-%d")
    today = date.today()
    trading_days = max((today - target_date).days * 5 // 7, 0)
    page = max(1, trading_days // PRICE_PAGE_SIZE)

    for _ in range(30):
        rows = _fetch_price_page(code, page)
        if not rows:
            return None
        newest, oldest = rows[0]["localTradedAt"], rows[-1]["localTradedAt"]
        for row in rows:
            if row["localTradedAt"] == date_str:
                try:
                    return int(str(row["closePrice"]).replace(",", ""))
                except (ValueError, TypeError):
                    return None
        if date_str > newest:
            if page == 1:
                return None
            page = max(1, page - 1)
        elif date_str < oldest:
            page += 1
        else:
            return None  # 범위 안이지만 해당일 없음 = 휴장일
    return None


def find_prev_trading_day(code: str, target_date: date) -> tuple[date | None, int | None]:
    for delta in range(0, 6):
        check_date = target_date - timedelta(days=delta)
        price = fetch_closing_price(code, check_date)
        if price:
            return check_date, price
    return None, None


# ── 스크린샷 (네이버 시세 페이지) ────────────────────────────
# finance.naver.com/item/* 은 전부 stock.naver.com으로 이전되어 폐지됨.
# m.stock.naver.com 페이지가 요약+차트+일별시세 표를 한 화면에 제공하며,
# 넓은 뷰포트로 열면 PC 화면처럼 표가 넓게 펼쳐져 표시됨(같은 페이지, 반응형 레이아웃).
CAPTURE_WIDTH = 1280
# 뷰포트를 넉넉히 크게 잡아 실제 스크롤이 일어나지 않게 함.
# (긴 목록을 스크롤하면 화면 밖으로 나간 행이 가상화(virtualization)로 DOM에서
#  사라져 캡처에서 잘려나가는 문제가 있어, 스크롤 자체를 피하는 방식으로 우회)
CAPTURE_VIEWPORT = {"width": CAPTURE_WIDTH, "height": 15000}


def capture_naver_chart(code: str, actual_date: date) -> tuple[dict, str]:
    """m.stock.naver.com 종목 시세 페이지를 PC 화면 크기로 그대로 열고, '더보기' 버튼만
    클릭해 기준일 행이 로드될 때까지 내려간 뒤 캡처. 페이지 구조/내용 수정 없음.
    """
    from playwright.sync_api import sync_playwright
    target_str = f"{actual_date.month:02d}. {actual_date.day:02d}."
    url = f"https://m.stock.naver.com/domestic/stock/{code}/price"
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
            page = browser.new_page(viewport=CAPTURE_VIEWPORT)
            page.goto(url, wait_until="networkidle", timeout=20000)
            page.wait_for_timeout(1000)

            found = False
            row_bottom = None
            for _ in range(40):
                loc = page.get_by_text(target_str, exact=True)
                if loc.count() > 0:
                    row_bottom = loc.first.evaluate(
                        "el => el.getBoundingClientRect().bottom + window.scrollY"
                    )
                    found = True
                    break
                more = page.query_selector("text=더보기")
                if not more:
                    break
                more.click()
                page.wait_for_timeout(400)

            if found and row_bottom:
                clip_h = int(row_bottom) + 40
                img = page.screenshot(clip={"x": 0, "y": 0, "width": CAPTURE_WIDTH, "height": clip_h})
                err = ""
            else:
                img = page.screenshot(clip={"x": 0, "y": 0, "width": CAPTURE_WIDTH, "height": 1200})
                err = "기준일 행을 찾지 못해 상단 화면만 캡처됨"

            browser.close()
            return {"전체": img}, err
    except Exception as e:
        return {}, str(e)[:200]


# ── 엑셀 생성 (다운로드용 메모리 버퍼) ──────────────────────
def save_excel(rows: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "종가현황"

    headers = ["펀드명", "종목명", "기준일", "실제 조회일", "종가(원)"]
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for r, row in enumerate(rows, 2):
        ws.cell(r, 1, row["fund"]).border = border
        ws.cell(r, 2, row["name"]).border = border
        ws.cell(r, 3, row["ref_date"]).border = border
        actual = ws.cell(r, 4, row["actual_date"])
        actual.border = border
        if row["ref_date"] != row["actual_date"]:
            actual.font = Font(color="FF0000")
        price_cell = ws.cell(r, 5, row["price"])
        price_cell.number_format = "#,##0"
        price_cell.alignment = Alignment(horizontal="right")
        price_cell.border = border

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────
st.markdown("### 1. 기준일 / 종목 리스트")
target_date = st.date_input("기준일", value=date.today())

mode = st.radio(
    "종목 리스트 방식",
    ["투자명부 자동 탐색 (로컬 전용)", "투자명부 파일 업로드 (원격/공유용)", "엑셀 직접 업로드 (펀드명/종목명)"],
    horizontal=True,
)

df = None
col_fund, col_name = "fund", "name"

if mode.startswith("투자명부 자동 탐색"):
    st.caption(f"경로: `{LEDGER_DIR}`")
    if st.button("투자명부 불러오기"):
        with st.spinner("전종목 코드 확인 및 투자명부 분석 중..."):
            code_map = get_code_map()
            ledger_path, ledger_label = find_ledger_file(target_date)
            if not ledger_path:
                st.session_state.pop("auto_df", None)
                st.error(f"{target_date:%Y년 %m월} 인근 투자명부 파일을 찾지 못했습니다. 경로/파일명을 확인해주세요.")
            else:
                auto_df, notices = extract_fund_stock_list(ledger_path, code_map)
                st.session_state["auto_df"] = auto_df
                st.session_state["auto_notices"] = notices
                st.session_state["auto_ledger_label"] = ledger_label

    if "auto_df" in st.session_state:
        st.caption(f"사용된 투자명부: **{st.session_state['auto_ledger_label']}**")
        for n in st.session_state.get("auto_notices", []):
            st.warning(n)
        df = st.session_state["auto_df"]

elif mode.startswith("투자명부 파일 업로드"):
    st.caption("투자명부 원본 파일(`투자명부(YYYY년MM월말기준).xlsx`)을 수정 없이 그대로 업로드하세요. "
               "PC/드라이브 접근 없이도(예: 다른 동료 PC) 같은 자동 리스트업 기능을 쓸 수 있습니다.")
    uploaded_ledger = st.file_uploader("투자명부 엑셀 (.xlsx)", type=["xlsx"], key="ledger_upload")
    if uploaded_ledger and st.button("투자명부 분석하기"):
        with st.spinner("전종목 코드 확인 및 투자명부 분석 중..."):
            code_map = get_code_map()
            auto_df, notices = extract_fund_stock_list(uploaded_ledger, code_map)
            st.session_state["auto_df"] = auto_df
            st.session_state["auto_notices"] = notices
            st.session_state["auto_ledger_label"] = f"업로드 파일: {uploaded_ledger.name}"

    if "auto_df" in st.session_state:
        st.caption(f"사용된 투자명부: **{st.session_state['auto_ledger_label']}**")
        for n in st.session_state.get("auto_notices", []):
            st.warning(n)
        df = st.session_state["auto_df"]

else:
    st.caption("**펀드명**, **종목명** 두 열 포함. 헤더 행 필수.")
    uploaded = st.file_uploader("종목 목록 엑셀 (.xlsx)", type=["xlsx"])
    if uploaded:
        try:
            raw_df = pd.read_excel(uploaded)
            raw_df.columns = raw_df.columns.str.strip()
            up_col_fund = next((c for c in raw_df.columns if "펀드" in str(c)), None)
            up_col_name = next((c for c in raw_df.columns if "종목" in str(c)), None)
            if not up_col_fund or not up_col_name:
                st.error("엑셀에 '펀드명'과 '종목명' 열이 필요합니다.")
            else:
                df = raw_df[[up_col_fund, up_col_name]].dropna()
                df.columns = [col_fund, col_name]
        except Exception as e:
            st.error(f"오류: {e}")
            import traceback
            st.code(traceback.format_exc())

capture_enabled = st.checkbox("📸 네이버 시세 화면 캡처 (분기결산 증빙용)", value=False,
                              help="체크 시 각 종목별 네이버 화면을 PNG로 저장. 처리 시간이 길어집니다.")
st.caption("⚠️ 기준일이 휴장일이면 직전 영업일 종가로 자동 대체됩니다. (엑셀에 빨간색 표시)")

if df is not None and not df.empty:
    try:
        st.dataframe(df, use_container_width=True)
        st.info(f"총 {len(df)}개 종목")

        if st.button("종가 조회 시작", type="primary"):
            code_map = get_code_map()
            if not code_map:
                st.error("종목코드 로딩 실패. 잠시 후 다시 시도해주세요.")
                st.stop()
            init_msgs = []
            if capture_enabled:
                with st.spinner("브라우저 초기화 중 (최초 1회 1~2분)..."):
                    ok, init_msgs = ensure_playwright_browser()
                if not ok:
                    st.error("브라우저 초기화 실패. 캡처 없이 종가 조회만 진행합니다.")
                    capture_enabled = False
            results = []
            errors = []
            captures: dict[str, bytes] = {}
            capture_cache: dict[str, tuple[dict, str]] = {}
            progress = st.progress(0)
            status = st.empty()

            for i, (_, row) in enumerate(df.iterrows()):
                fund_name = str(row[col_fund]).strip()
                stock_name = str(row[col_name]).strip()
                status.text(f"처리 중: {stock_name} ({i+1}/{len(df)})")

                code = get_stock_code(stock_name, code_map)
                if not code:
                    errors.append(f"{stock_name}: 종목코드 없음")
                    results.append({
                        "fund": fund_name, "name": stock_name,
                        "ref_date": target_date.strftime("%Y-%m-%d"),
                        "actual_date": "-", "price": "코드 없음",
                    })
                    progress.progress((i + 1) / len(df))
                    continue

                actual_date, price = find_prev_trading_day(code, target_date)

                if not price:
                    errors.append(f"{stock_name}: 종가 조회 실패")
                    results.append({
                        "fund": fund_name, "name": stock_name,
                        "ref_date": target_date.strftime("%Y-%m-%d"),
                        "actual_date": "-", "price": "조회 실패",
                    })
                    progress.progress((i + 1) / len(df))
                    continue

                results.append({
                    "fund": fund_name, "name": stock_name,
                    "ref_date": target_date.strftime("%Y-%m-%d"),
                    "actual_date": actual_date.strftime("%Y-%m-%d"),
                    "price": price,
                })

                if capture_enabled:
                    status.text(f"캡처 중: {stock_name} ({i+1}/{len(df)})")
                    cache_key = f"{code}_{actual_date.isoformat()}"
                    if cache_key in capture_cache:
                        imgs, err = capture_cache[cache_key]
                    else:
                        imgs, err = capture_naver_chart(code, actual_date)
                        capture_cache[cache_key] = (imgs, err)
                    date_label = target_date.strftime("%Y%m%d")
                    for suffix, data in imgs.items():
                        captures[f"{fund_name}/{stock_name}_{date_label}_{suffix}.png"] = data
                    if not imgs:
                        errors.append(f"{stock_name}: 캡처 실패 — {err}")
                    elif err:
                        errors.append(f"{stock_name}: 일부 캡처 실패 — {err}")

                progress.progress((i + 1) / len(df))

            status.text("완료!")

            excel_rows = [r for r in results if isinstance(r["price"], int)]
            date_label = target_date.strftime("%Y%m%d")
            excel_bytes = save_excel(excel_rows) if excel_rows else None

            zip_bytes = None
            if captures:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname, data in captures.items():
                        zf.writestr(fname, data)
                zip_bytes = zip_buf.getvalue()

            # session_state에 저장 → 다운로드 클릭 후에도 결과 유지
            st.session_state["last_results"] = results
            st.session_state["last_errors"] = errors
            st.session_state["last_excel"] = excel_bytes
            st.session_state["last_zip"] = zip_bytes
            st.session_state["last_zip_count"] = len(captures)
            st.session_state["last_date_label"] = date_label
            st.session_state["last_init_msgs"] = init_msgs

    except Exception as e:
        st.error(f"오류: {e}")
        import traceback
        st.code(traceback.format_exc())

# ── 결과 표시 (session_state 기반, 다운로드해도 유지됨) ──────
if "last_results" in st.session_state:
    st.markdown("---")
    st.markdown("### 결과")
    result_df = pd.DataFrame(st.session_state["last_results"])
    result_df.columns = ["펀드명", "종목명", "기준일", "실제조회일", "종가(원)"]
    st.dataframe(result_df, use_container_width=True)

    col_a, col_b = st.columns(2)
    date_label = st.session_state["last_date_label"]
    if st.session_state.get("last_excel"):
        with col_a:
            st.download_button(
                label="📥 엑셀 다운로드",
                data=st.session_state["last_excel"],
                file_name=f"종가현황_{date_label}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_excel",
            )
    if st.session_state.get("last_zip"):
        with col_b:
            st.download_button(
                label=f"📸 캡처 ZIP 다운로드 ({st.session_state['last_zip_count']}장)",
                data=st.session_state["last_zip"],
                file_name=f"종가캡처_{date_label}.zip",
                mime="application/zip",
                key="dl_zip",
            )

    if st.session_state.get("last_init_msgs"):
        with st.expander("브라우저 초기화 로그"):
            for m in st.session_state["last_init_msgs"]:
                st.text(m)

    if st.session_state.get("last_errors"):
        with st.expander("오류 목록"):
            for e in st.session_state["last_errors"]:
                st.warning(e)
