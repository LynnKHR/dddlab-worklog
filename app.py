# app.py
# 디디디랩 근무관리
# 관리자 페이지 + 직원 토큰 페이지 (Streamlit Secrets 버전)

from __future__ import annotations
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

# =========================
# 설정
# =========================
SPREADSHEET_KEY = "1kYxMXFc78uBW4tZpNelLJtaKlvkUlwZC31CWBfEz0mE"
PENDING_SHEET_NAME = "승인대기"
EMPLOYEE_SHEET_NAME = "직원목록"

FINAL_HEADER = ["이름", "날짜", "출근시간", "퇴근시간", "근무시간(시)"]
PENDING_HEADER = ["이름", "날짜", "출근시간", "퇴근시간", "승인", "승인시각"]

CACHE_TTL_SECONDS = 30

# =========================
# 시간 검증
# =========================
TIME_PATTERN = re.compile(r"^\d{1,2}:\d{1,2}$")

def parse_and_format_time(s: str):
    if not isinstance(s, str):
        return None
    if not TIME_PATTERN.match(s.strip()):
        return None
    h, m = map(int, s.split(":"))
    if not (0 <= h <= 23 and 0 <= m <= 59):
        return None
    return f"{h:02d}:{m:02d}"

def is_in_work_range(t: str) -> bool:
    return 6 <= int(t[:2]) <= 22

def validate_times(in_t, out_t):
    if in_t:
        ft = parse_and_format_time(in_t)
        if not ft or not is_in_work_range(ft):
            return False, "출근시간은 06:00~22:00 HH:MM 형식이어야 합니다."
    if out_t:
        ft = parse_and_format_time(out_t)
        if not ft or not is_in_work_range(ft):
            return False, "퇴근시간은 06:00~22:00 HH:MM 형식이어야 합니다."
    if in_t and out_t and parse_and_format_time(in_t) >= parse_and_format_time(out_t):
        return False, "출근시간은 퇴근시간보다 빨라야 합니다."
    return True, ""

def now_hhmm():
    return datetime.now().strftime("%H:%M")

def today():
    return datetime.now().strftime("%Y-%m-%d")

def calc_hours(i, o):
    if not i or not o:
        return ""
    t1 = datetime.strptime(i, "%H:%M")
    t2 = datetime.strptime(o, "%H:%M")
    return round((t2 - t1).total_seconds() / 3600, 2)

# =========================
# Google Sheets (Secrets)
# =========================
@st.cache_resource
def get_client():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        st.secrets["gcp_service_account"],
        scope
    )
    return gspread.authorize(creds)

@st.cache_resource
def get_spreadsheet():
    return get_client().open_by_key(SPREADSHEET_KEY)

@st.cache_resource
def get_worksheets():
    ss = get_spreadsheet()

    final_ws = ss.sheet1

    try:
        pending_ws = ss.worksheet(PENDING_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        pending_ws = ss.add_worksheet(PENDING_SHEET_NAME, 2000, 10)
        pending_ws.append_row(PENDING_HEADER)

    try:
        emp_ws = ss.worksheet(EMPLOYEE_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        emp_ws = ss.add_worksheet(EMPLOYEE_SHEET_NAME, 1000, 5)
        emp_ws.append_row(["이름", "토큰"])

    return final_ws, pending_ws, emp_ws

@st.cache_data(ttl=CACHE_TTL_SECONDS)
def read_df(sheet_name):
    ss = get_spreadsheet()
    ws = ss.sheet1 if sheet_name == "__FINAL__" else ss.worksheet(sheet_name)
    return pd.DataFrame(ws.get_all_records())

def clear_read_cache():
    read_df.clear()

# =========================
# 시트 유틸
# =========================
def find_row(ws, name, date):
    for i, r in enumerate(ws.get_all_values()[1:], start=2):
        if r[0] == name and r[1] == date:
            return i
    return None

def upsert_pending(ws, name, date, in_t=None, out_t=None):
    r = find_row(ws, name, date)
    if r:
        if in_t is not None:
            ws.update_cell(r, 3, in_t)
        if out_t is not None:
            ws.update_cell(r, 4, out_t)
    else:
        ws.append_row([name, date, in_t or "", out_t or "", False, ""])

def approve(final_ws, name, date, i, o):
    row = [name, date, i, o, calc_hours(i, o)]
    r = find_row(final_ws, name, date)
    if r:
        final_ws.update(f"A{r}:E{r}", [row])
    else:
        final_ws.append_row(row)

def revoke(final_ws, name, date):
    r = find_row(final_ws, name, date)
    if r:
        final_ws.delete_rows(r)

# =========================
# 🔐 토큰 판별
# =========================
query = st.query_params
token = query.get("token", [None])[0]

current_user = None
is_employee_page = False

if token:
    emp_df = read_df(EMPLOYEE_SHEET_NAME)
    row = emp_df[emp_df["토큰"] == token]
    if row.empty:
        st.error("유효하지 않은 접근입니다.")
        st.stop()
    current_user = row.iloc[0]["이름"]
    is_employee_page = True

# =========================
# UI
# =========================
st.set_page_config(layout="wide")
st.title("🕓 디디디랩 근무관리")

final_ws, pending_ws, emp_ws = get_worksheets()
pending_df = read_df(PENDING_SHEET_NAME)
final_df = read_df("__FINAL__")

if not pending_df.empty:
    pending_df["승인"] = pending_df["승인"].astype(str).str.lower().isin(["true", "1", "yes"])
    pending_df["승인시각"] = pd.to_datetime(pending_df["승인시각"], errors="coerce")

# =========================
# 👤 직원 페이지
# =========================
if is_employee_page:
    st.info(f"👤 사용자: {current_user}")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("출근"):
            t = now_hhmm()
            ok, msg = validate_times(t, None)
            if ok:
                upsert_pending(pending_ws, current_user, today(), in_t=t)
                clear_read_cache()
                st.success("출근 등록 완료")
                st.rerun()
            else:
                st.error(msg)

    with c2:
        if st.button("퇴근"):
            t = now_hhmm()
            ok, msg = validate_times(None, t)
            if ok:
                upsert_pending(pending_ws, current_user, today(), out_t=t)
                clear_read_cache()
                st.success("퇴근 등록 완료")
                st.rerun()
            else:
                st.error(msg)

    st.divider()
    st.subheader("📋 내 근무 기록")
    st.dataframe(pending_df[pending_df["이름"] == current_user], use_container_width=True)
    st.stop()

# =========================
# 🛡 관리자 페이지
# =========================
st.subheader("🛡 관리자 페이지")
st.dataframe(pending_df, use_container_width=True)

st.divider()
st.subheader("📘 확정 근무 기록")
st.dataframe(final_df, use_container_width=True)
