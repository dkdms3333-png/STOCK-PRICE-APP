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

st.set_page_config(page_title="종가 캡처", layout="centered")
st.title("코스닥/코스피 종가 조회")

# ── 종목코드 조회 ─────────────────────────────────────────────
def clean_stock_name(name: str) -> str:
    return re.sub(r'[㈜㈔\s]', '', name).strip()


_code_cache: dict[str, str] = {}

def get_stock_code(name: str, code_map: dict = None) -> str | None:
    """네이버 금융 자동완성 API로 종목명 → 코드 반환."""
    clean = clean_stock_name(name.strip())

    if clean in _code_cache:
        return _code_cache[clean]

    try:
        url = f"https://ac.finance.naver.com/ac?q={requests.utils.quote(clean)}&q_enc=UTF-8&t_koreng=1&st=111&r_format=json&r_enc=UTF-8&r_lt=111&r_unicode=0&r_escape=1"
        headers = {"User-Agent": "Mozilla/5.0"}
        res = requests.get(url, headers=headers, timeout=5)
        data = res.json()
        items = data.get("items", [[]])[0]
        if items:
            code = items[0][1] if len(items[0]) > 1 else None
            if code:
                _code_cache[clean] = code
                return code
    except Exception:
        pass
    return None


# ── 종가 조회 ────────────────────────────────────────────────
def fetch_closing_price(code: str, target_date: date) -> int | None:
    headers = {"User-Agent": "Mozilla/5.0"}
    date_str = target_date.strftime("%Y.%m.%d")
    today = date.today()
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
            if row_date and row_date < date_str:
                break
    return None


def find_prev_trading_day(code: str, target_date: date) -> tuple[date | None, int | None]:
    for delta in range(0, 6):
        check_date = target_date - timedelta(days=delta)
        price = fetch_closing_price(code, check_date)
        if price:
            return check_date, price
    return None, None


# ── 엑셀 생성 (메모리 버퍼) ──────────────────────────────────
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
st.markdown("### 1. 엑셀 업로드")
st.caption("**펀드명**, **종목명** 두 열 포함. 헤더 행 필수.")

uploaded = st.file_uploader("종목 목록 엑셀 (.xlsx)", type=["xlsx"])
target_date = st.date_input("기준일", value=date.today())
st.caption("⚠️ 기준일이 휴장일이면 직전 영업일 종가로 자동 대체됩니다. (엑셀에 빨간색 표시)")

if uploaded:
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

            if st.button("종가 조회 시작", type="primary"):
                results = []
                errors = []
                progress = st.progress(0)
                status = st.empty()

                for i, (_, row) in enumerate(df.iterrows()):
                    fund_name = str(row[col_fund]).strip()
                    stock_name = str(row[col_name]).strip()
                    status.text(f"처리 중: {stock_name} ({i+1}/{len(df)})")

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
                    progress.progress((i + 1) / len(df))

                status.text("완료!")

                result_df = pd.DataFrame(results)
                result_df.columns = ["펀드명", "종목명", "기준일", "실제조회일", "종가(원)"]
                st.dataframe(result_df, use_container_width=True)

                # 엑셀 다운로드
                excel_rows = [r for r in results if isinstance(r["price"], int)]
                if excel_rows:
                    date_label = target_date.strftime("%Y%m%d")
                    excel_bytes = save_excel(excel_rows)
                    st.download_button(
                        label="📥 엑셀 다운로드",
                        data=excel_bytes,
                        file_name=f"종가현황_{date_label}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                if errors:
                    with st.expander("오류 목록"):
                        for e in errors:
                            st.warning(e)

    except Exception as e:
        st.error(f"오류: {e}")
        import traceback
        st.code(traceback.format_exc())
