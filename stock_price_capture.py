import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from playwright.sync_api import sync_playwright
from datetime import datetime, date, timedelta
import requests
from bs4 import BeautifulSoup
import re
import os
import time
from pathlib import Path

st.set_page_config(page_title="종가 캡처", layout="centered")
st.title("코스닥/코스피 종가 조회 및 캡처")

# ── 종목코드 조회 ─────────────────────────────────────────────
def clean_stock_name(name: str) -> str:
    """㈜, ㈔ 등 법인 기호·공백 제거"""
    return re.sub(r'[㈜㈔\s]', '', name).strip()


# 세션 내 코드 캐시 (Playwright 검색 중복 방지)
_code_cache: dict[str, str] = {}

def get_stock_code(name: str, code_map: dict = None) -> str | None:
    """네이버 금융 검색으로 종목명 → 코드 반환.
    세션 캐시 → Playwright 검색 순으로 시도.
    최근 상장 종목 포함 모든 종목 커버.
    """
    clean = clean_stock_name(name.strip())

    # 캐시 확인
    if clean in _code_cache:
        return _code_cache[clean]

    # Playwright로 네이버 금융 검색창에 직접 입력
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto("https://finance.naver.com/", wait_until="networkidle", timeout=15000)
            time.sleep(0.5)

            search_box = page.query_selector("input.search_input, input#stock_items, input[placeholder*='종목']")
            if search_box:
                search_box.click()
                search_box.fill(clean)
                time.sleep(0.8)
                search_box.press("Enter")
                try:
                    page.wait_for_url("**/item/**", timeout=5000)
                except:
                    pass
                time.sleep(1)
                url = page.url
                m = re.search(r'code=([\w]+)', url)
                if m:
                    code = m.group(1)
                    _code_cache[clean] = code
                    browser.close()
                    return code
            browser.close()
    except Exception:
        pass
    return None


# ── 종가 조회 (sise_day 직접 스크래핑) ───────────────────────
def fetch_closing_price(code: str, target_date: date) -> int | None:
    """네이버 금융 일별시세에서 특정일 종가 반환."""
    headers = {"User-Agent": "Mozilla/5.0"}
    date_str = target_date.strftime("%Y.%m.%d")
    today = date.today()
    # 영업일 기준 페이지 추정 (1페이지 = 약 10영업일)
    trading_days = (today - target_date).days * 5 // 7
    start_page = max(1, trading_days // 10 - 2)

    for page in range(start_page, start_page + 15):
        url = f"https://finance.naver.com/item/sise_day.nhn?code={code}&page={page}"
        res = requests.get(url, headers=headers, timeout=10)
        res.encoding = "euc-kr"
        soup = BeautifulSoup(res.text, "html.parser")
        rows = soup.select("table.type2 tr")
        for row in rows:
            tds = row.select("td")
            if len(tds) < 2:
                continue
            row_date = tds[0].text.strip()
            if row_date == date_str:
                price_text = tds[1].text.strip().replace(",", "")
                try:
                    return int(price_text)
                except ValueError:
                    return None
            # 현재 행이 기준일보다 과거면 다음 페이지로
            if row_date and row_date < date_str:
                break
    return None


def find_prev_trading_day(code: str, target_date: date, code_map: dict) -> tuple[date | None, int | None]:
    """기준일에 데이터 없으면 직전 영업일 탐색 (최대 5일)"""
    for delta in range(0, 6):
        check_date = target_date - timedelta(days=delta)
        price = fetch_closing_price(code, check_date)
        if price:
            return check_date, price
    return None, None


# ── 스크린샷 (네이버 금융 시세 페이지) ──────────────────────
def find_page_for_date(code: str, target_date: date) -> int:
    """sise_day에서 해당 날짜가 있는 정확한 페이지 번호 반환"""
    headers = {"User-Agent": "Mozilla/5.0"}
    date_str = target_date.strftime("%Y.%m.%d")
    today = date.today()
    trading_days = (today - target_date).days * 5 // 7
    start_page = max(1, trading_days // 10 - 2)
    for page in range(start_page, start_page + 10):
        url = f"https://finance.naver.com/item/sise_day.nhn?code={code}&page={page}"
        res = requests.get(url, headers=headers, timeout=10)
        res.encoding = "euc-kr"
        soup = BeautifulSoup(res.text, "html.parser")
        for row in soup.select("table.type2 tr"):
            tds = row.select("td")
            if tds and tds[0].text.strip() == date_str:
                return page
            if tds and tds[0].text.strip() and tds[0].text.strip() < date_str:
                break
    return start_page


def capture_naver_chart(code: str, stock_name: str, actual_date: date, save_path: str):
    """finance.naver.com/item/sise.naver 캡처.
    종목명(상단) + 기준일 종가가 있는 일별시세 페이지를 'day' 프레임에 로드.
    """
    target_page = find_page_for_date(code, actual_date)
    sise_url = f"https://finance.naver.com/item/sise.naver?code={code}"
    day_url  = f"https://finance.naver.com/item/sise_day.naver?code={code}&page={target_page}"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        pw_page = browser.new_page(viewport={"width": 1280, "height": 1300})

        pw_page.goto(sise_url, wait_until="networkidle", timeout=30000)
        time.sleep(1.5)

        day_frame = pw_page.frame(name="day")
        if day_frame:
            day_frame.goto(day_url, wait_until="networkidle", timeout=15000)
            time.sleep(1)

        pw_page.screenshot(path=save_path, clip={"x": 0, "y": 0, "width": 1280, "height": 1300})
        browser.close()


# ── 엑셀 저장 ─────────────────────────────────────────────────
def save_excel(rows: list[dict], save_path: str):
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
            actual.font = Font(color="FF0000")  # 빨간색: 직전 영업일 사용
        price_cell = ws.cell(r, 5, row["price"])
        price_cell.number_format = "#,##0"
        price_cell.alignment = Alignment(horizontal="right")
        price_cell.border = border

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14

    wb.save(save_path)


# ── UI ────────────────────────────────────────────────────────
st.markdown("### 1. 엑셀 업로드")
st.caption("**펀드명**, **종목명** 두 열 포함. 헤더 행 필수.")

uploaded = st.file_uploader("종목 목록 엑셀 (.xlsx)", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    target_date = st.date_input("기준일", value=date.today())
with col2:
    save_folder = st.text_input(
        "저장 폴더",
        value=r"C:\Users\최아은\.claude\ani\종가조회",
    )

st.caption("⚠️ 기준일이 휴장일이면 직전 영업일 종가로 자동 대체됩니다. (엑셀에 빨간색 표시)")

if uploaded and save_folder:
    try:
        df = pd.read_excel(uploaded)
        df.columns = df.columns.str.strip()

        col_fund = next((c for c in df.columns if "펀드" in str(c)), None)
        col_name = next((c for c in df.columns if "종목" in str(c)), None)

        if not col_fund or not col_name:
            st.error("엑셀에 '펀드명'과 '종목명' 열이 필요합니다.")
        else:
            df = df[[col_fund, col_name]].dropna()
            st.dataframe(df, use_container_width=True)
            st.info(f"총 {len(df)}개 종목")

            if st.button("종가 조회 + 캡처 시작", type="primary"):
                save_dir = Path(save_folder)
                save_dir.mkdir(parents=True, exist_ok=True)

                date_label = target_date.strftime("%Y%m%d")
                results = []
                errors = []
                progress = st.progress(0)
                status = st.empty()

                for i, (_, row) in enumerate(df.iterrows()):
                    fund_name = str(row[col_fund]).strip()
                    stock_name = str(row[col_name]).strip()
                    status.text(f"처리 중: {stock_name} ({i+1}/{len(df)}) — 종목코드 검색 중...")

                    code = get_stock_code(stock_name)
                    if not code:
                        errors.append(f"{stock_name}: 종목코드 없음")
                        results.append({
                            "fund": fund_name, "name": stock_name,
                            "ref_date": target_date.strftime("%Y-%m-%d"),
                            "actual_date": "-", "price": "코드 없음",
                        })
                        progress.progress((i + 1) / len(df))
                        continue

                    actual_date, price = find_prev_trading_day(code, target_date, {})

                    if not price:
                        errors.append(f"{stock_name}: 종가 조회 실패")
                        results.append({
                            "fund": fund_name, "name": stock_name,
                            "ref_date": target_date.strftime("%Y-%m-%d"),
                            "actual_date": "-", "price": "조회 실패",
                        })
                        progress.progress((i + 1) / len(df))
                        continue

                    # 스크린샷
                    img_filename = f"{stock_name}_{date_label}.png"
                    img_path = str(save_dir / img_filename)
                    try:
                        capture_naver_chart(code, stock_name, actual_date, img_path)
                    except Exception as e:
                        errors.append(f"{stock_name} 캡처 실패: {e}")

                    results.append({
                        "fund": fund_name, "name": stock_name,
                        "ref_date": target_date.strftime("%Y-%m-%d"),
                        "actual_date": actual_date.strftime("%Y-%m-%d"),
                        "price": price,
                    })
                    progress.progress((i + 1) / len(df))

                # 엑셀 저장
                excel_rows = [r for r in results if isinstance(r["price"], int)]
                if excel_rows:
                    excel_path = str(save_dir / f"종가현황_{date_label}.xlsx")
                    save_excel(excel_rows, excel_path)

                status.text("완료!")
                st.success(f"저장 완료 → {save_folder}")

                result_df = pd.DataFrame(results)
                result_df.columns = ["펀드명", "종목명", "기준일", "실제조회일", "종가(원)"]
                st.dataframe(result_df, use_container_width=True)

                if errors:
                    with st.expander("오류 목록"):
                        for e in errors:
                            st.warning(e)

    except Exception as e:
        st.error(f"오류: {e}")
        import traceback
        st.code(traceback.format_exc())
