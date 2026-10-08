"""
Route Coordinate Matcher — Streamlit Web App
==============================================
โยนไฟล์ route เข้าไป ได้พิกัดกลับมาทันที ผ่านเว็บแอปที่มีลิงก์ถาวร

วิธี deploy: ดู DEPLOY_INSTRUCTIONS.md
"""

import streamlit as st
import pandas as pd
import io
import re
import xlrd
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
import qrcode

st.set_page_config(page_title="Route Coordinate Matcher", page_icon="🚚", layout="wide")

# ============================================================
# THEME — Logistics / Transportation, refined
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .stApp {
        background:
            radial-gradient(1200px 600px at 90% -10%, rgba(255,107,53,0.08), transparent 60%),
            radial-gradient(1000px 500px at -10% 10%, rgba(47,128,237,0.10), transparent 60%),
            linear-gradient(180deg, #070C16 0%, #0B1524 45%, #0A1220 100%);
    }

    /* ---------- Header (with photo) ---------- */
    .rcm-header {
        background:
            linear-gradient(100deg, rgba(9,17,32,0.94) 0%, rgba(11,22,40,0.85) 40%, rgba(11,22,40,0.35) 75%, rgba(11,22,40,0.15) 100%),
            url('https://images.unsplash.com/photo-1601584115197-04ecc0da31d7?auto=format&fit=crop&w=1600&q=80');
        background-size: cover;
        background-position: center 60%;
        border: 1px solid #204A78;
        border-radius: 18px;
        padding: 30px 34px 24px 34px;
        margin-bottom: 22px;
        position: relative;
        overflow: hidden;
        box-shadow: 0 10px 40px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.04);
        min-height: 210px;
    }
    .rcm-header::before {
        content: "";
        position: absolute;
        left: 0; bottom: 0; width: 100%; height: 3px;
        background: linear-gradient(90deg, #FF6B35 0%, #FFB347 35%, transparent 75%);
    }
    .rcm-eyebrow {
        font-family: 'JetBrains Mono', monospace;
        font-size: 11.5px;
        letter-spacing: 0.18em;
        color: #FF9B5C;
        font-weight: 700;
        margin-bottom: 8px;
        text-transform: uppercase;
        display: flex; align-items: center; gap: 6px;
    }
    .rcm-header-content {
        position: relative;
        z-index: 2;
    }
    .rcm-route-svg {
        position: absolute;
        top: 0; right: 0;
        width: 60%; height: 100%;
        z-index: 1;
        pointer-events: none;
    }
    .rcm-title {
        font-size: 30px;
        font-weight: 800;
        color: #F7FAFD;
        margin: 0 0 6px 0;
        letter-spacing: -0.02em;
        text-shadow: 0 2px 12px rgba(0,0,0,0.5);
    }
    .rcm-subtitle {
        font-size: 14px;
        color: #C3D4EA;
        font-family: 'JetBrains Mono', monospace;
        margin-bottom: 20px;
    }
    .rcm-features {
        display: flex;
        gap: 28px;
        flex-wrap: wrap;
    }
    .rcm-feature {
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .rcm-feature-icon {
        width: 30px; height: 30px;
        border-radius: 8px;
        background: rgba(255,255,255,0.08);
        border: 1px solid rgba(255,255,255,0.14);
        display: flex; align-items: center; justify-content: center;
        font-size: 15px;
        flex-shrink: 0;
    }
    .rcm-feature-text {
        line-height: 1.25;
    }
    .rcm-feature-title {
        font-size: 12.5px;
        font-weight: 700;
        color: #F0F5FA;
    }
    .rcm-feature-sub {
        font-size: 10.5px;
        color: #92A9C4;
        font-family: 'JetBrains Mono', monospace;
    }

    /* ---------- Status / connection card ---------- */
    .rcm-status-card {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: linear-gradient(90deg, rgba(42,110,76,0.20), rgba(12,23,40,0.55));
        border: 1px solid #2A6E4C;
        border-radius: 14px;
        padding: 16px 22px;
        margin-bottom: 14px;
        box-shadow: 0 4px 18px rgba(0,0,0,0.22);
    }
    .rcm-status-left { display: flex; align-items: center; gap: 14px; }
    .rcm-status-icon {
        width: 38px; height: 38px; border-radius: 10px;
        background: rgba(52,199,89,0.18);
        border: 1px solid rgba(52,199,89,0.4);
        display: flex; align-items: center; justify-content: center;
        font-size: 17px;
    }
    .rcm-status-title { font-weight: 700; color: #F0F5FA; font-size: 14.5px; }
    .rcm-status-sub { font-size: 11.5px; color: #8FA8C7; font-family: 'JetBrains Mono', monospace; }
    .rcm-status-right { display: flex; align-items: center; gap: 8px; }
    .rcm-status-right-icon {
        width: 34px; height: 34px; border-radius: 9px;
        background: rgba(255,255,255,0.06);
        border: 1px solid rgba(255,255,255,0.12);
        display: flex; align-items: center; justify-content: center;
        font-size: 15px;
    }
    .rcm-status-pill {
        font-family: 'JetBrains Mono', monospace;
        font-size: 11px; color: #6FE39A; font-weight: 700;
        display: flex; align-items: center; gap: 5px;
    }
    .rcm-status-dot {
        width: 6px; height: 6px; border-radius: 50%;
        background: #34C759;
        box-shadow: 0 0 6px #34C759;
    }
    .rcm-sync-btn-wrap { height: 100%; display: flex; align-items: center; }
    .rcm-sync-btn-wrap .stButton { width: 100%; }
    .rcm-sync-btn-wrap .stButton button {
        height: 100%;
        min-height: 66px;
        background: linear-gradient(180deg, #101E33, #0C1728) !important;
        border: 1px solid #1E3A5F !important;
        color: #C3D4EA !important;
        font-size: 12px !important;
        white-space: normal !important;
        line-height: 1.3;
    }
    .rcm-sync-btn-wrap .stButton button:hover {
        border-color: #FF8C42 !important;
        color: #F7FAFD !important;
    }

    /* Expander styled like clickable list card */
    div[data-testid="stExpander"] {
        border: 1px solid #1E3A5F !important;
        border-radius: 14px !important;
        background: linear-gradient(90deg, rgba(47,128,237,0.10), rgba(12,23,40,0.55)) !important;
        overflow: hidden;
    }
    div[data-testid="stExpander"] summary {
        padding: 14px 20px !important;
        font-weight: 600 !important;
    }

    /* ---------- Cards ---------- */
    .rcm-card {
        background: linear-gradient(180deg, #101E33 0%, #0C1728 100%);
        border: 1px solid #1E3A5F;
        border-radius: 14px;
        padding: 18px 22px;
        margin-bottom: 16px;
        box-shadow: 0 4px 18px rgba(0,0,0,0.22);
    }
    .rcm-card-success {
        border-color: #2A6E4C;
        background: linear-gradient(120deg, rgba(42,110,76,0.18), rgba(12,23,40,0.5));
    }
    .rcm-card-warn {
        border-color: #8A6329;
        background: linear-gradient(120deg, rgba(138,99,41,0.18), rgba(12,23,40,0.5));
    }

    /* ---------- Stat boxes ---------- */
    .rcm-stat {
        background: linear-gradient(180deg, #0F1D32 0%, #0A1526 100%);
        border: 1px solid #1E3A5F;
        border-radius: 12px;
        padding: 16px 18px;
        text-align: center;
        transition: border-color 0.18s ease, transform 0.18s ease, box-shadow 0.18s ease;
    }
    .rcm-stat:hover {
        border-color: #FF8C42;
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(255,107,53,0.15);
    }
    .rcm-stat-label {
        font-family: 'JetBrains Mono', monospace;
        font-size: 11px;
        color: #6B84A6;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        margin-bottom: 8px;
    }
    .rcm-stat-value {
        font-size: 24px;
        font-weight: 800;
        color: #F7FAFD;
        background: linear-gradient(90deg, #F7FAFD, #B8D4F0);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .rcm-stat-sub {
        font-family: 'JetBrains Mono', monospace;
        font-size: 11px;
        color: #FF9B5C;
        margin-top: 4px;
    }

    /* ---------- Buttons ---------- */
    .stDownloadButton button, .stButton button {
        border-radius: 10px !important;
        font-weight: 700 !important;
        border: none !important;
        transition: transform 0.15s ease, box-shadow 0.15s ease !important;
    }
    .stDownloadButton button {
        background: linear-gradient(90deg, #FF8C42, #FF6B35) !important;
        color: #0B1220 !important;
        box-shadow: 0 4px 16px rgba(255,107,53,0.28) !important;
    }
    .stDownloadButton button:hover, .stButton button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 22px rgba(255,107,53,0.4) !important;
    }
    button[kind="primary"] {
        background: linear-gradient(90deg, #2F80ED, #1E63C4) !important;
        color: #F7FAFD !important;
        box-shadow: 0 4px 16px rgba(47,128,237,0.3) !important;
    }

    /* Custom "preview" toggle button styled as a list card */
    .rcm-preview-toggle-wrap .stButton button {
        background: linear-gradient(90deg, rgba(47,128,237,0.10), rgba(12,23,40,0.55)) !important;
        border: 1px solid #1E3A5F !important;
        border-radius: 14px !important;
        color: #E8F0FA !important;
        text-align: left !important;
        justify-content: flex-start !important;
        padding: 16px 20px !important;
        font-weight: 500 !important;
        font-size: 13.5px !important;
        box-shadow: none !important;
    }
    .rcm-preview-toggle-wrap .stButton button:hover {
        border-color: #2F80ED !important;
        transform: none !important;
        box-shadow: 0 4px 14px rgba(47,128,237,0.18) !important;
    }

    [data-testid="stFileUploaderDropzone"] {
        background: linear-gradient(180deg, #0F1D32, #0A1526) !important;
        border: 2px dashed #2A4A72 !important;
        border-radius: 14px !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        background: linear-gradient(90deg, #2F80ED, #1E63C4) !important;
        border-radius: 8px !important;
        border: none !important;
        outline: none !important;
        box-shadow: none !important;
        position: relative !important;
        padding: 9px 20px !important;
        min-width: 150px;
        min-height: 38px;
    }
    [data-testid="stFileUploaderDropzone"] button:focus,
    [data-testid="stFileUploaderDropzone"] button:focus-visible {
        outline: none !important;
        box-shadow: none !important;
        border-color: transparent !important;
    }
    [data-testid="stFileUploaderDropzone"] button * {
        display: none !important;
    }
    [data-testid="stFileUploaderDropzone"] button::after {
        content: "↑  Upload File";
        position: absolute;
        inset: 0;
        display: flex !important;
        align-items: center;
        justify-content: center;
        font-size: 13.5px;
        font-weight: 600;
        color: #F7FAFD;
        font-family: 'Inter', sans-serif;
        white-space: nowrap;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] span {
        font-size: 0 !important;
        display: block;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] span::after {
        content: "ลากไฟล์ route มาวางที่นี่ หรือคลิกเพื่อเลือกไฟล์";
        font-size: 14px;
        font-weight: 600;
        color: #E8F0FA;
        font-family: 'Inter', sans-serif;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] small {
        color: #6B84A6 !important;
        font-family: 'JetBrains Mono', monospace !important;
    }

    /* ---------- Dataframe ---------- */
    [data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; border: 1px solid #1E3A5F; }

    hr { border-color: #1E3A5F !important; }

    /* Radio pills for export format */
    div[role="radiogroup"] { gap: 8px; }

    /* subtle divider glow */
    .rcm-glow-divider {
        height: 1px;
        background: linear-gradient(90deg, transparent, #2A4A72, transparent);
        margin: 22px 0;
    }
    /* ---------- Tabs: ตัวหนังสือสว่างตลอด ไม่ต้องกดก่อนถึงจะเห็น ---------- */
    [data-testid="stTab"], [role="tab"], button[data-baseweb="tab"] { color: #F0F5FA !important; }
    [data-testid="stTab"] p, [role="tab"] p, button[data-baseweb="tab"] p {
        color: #F0F5FA !important; font-weight: 600; font-size: 15px;
    }
    [data-testid="stTab"]:hover p, [role="tab"]:hover p, button[data-baseweb="tab"]:hover p { color: #FFB347 !important; }
    [data-testid="stTab"][aria-selected="true"] p, [role="tab"][aria-selected="true"] p,
    button[data-baseweb="tab"][aria-selected="true"] p { color: #FF8C42 !important; font-weight: 700; }
    [role="tablist"], [data-baseweb="tab-list"] { border-bottom: 1px solid rgba(195,212,234,0.18); }
    [data-baseweb="tab-highlight"], .react-aria-SelectionIndicator { background-color: #FF8C42 !important; }
</style>
""", unsafe_allow_html=True)

MASTER_SHEET_CSV_URL = st.secrets.get("MASTER_SHEET_CSV_URL", "")
PRODUCT_SHEET_CSV_URL = st.secrets.get("PRODUCT_SHEET_CSV_URL", "")


# ============================================================
# HELPERS
# ============================================================
def normalize(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ''
    return re.sub(r'\s+', ' ', str(s).strip().upper().replace('.', '').replace(',', ''))


@st.cache_data(ttl=300)
def load_master_data(url):
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    required = ['Ship To', 'Ship To Name', 'Latitude', 'Longitude']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"ไม่พบคอลัมน์: {', '.join(missing)} กรุณาตรวจสอบหัวตารางใน Google Sheets")
    df = df.dropna(subset=['Ship To', 'Latitude', 'Longitude'])
    df['norm_ship'] = df['Ship To Name'].apply(normalize)
    df['norm_code'] = df['Ship To'].astype(str).str.strip().str.upper()
    return df


@st.cache_data(ttl=300)
def load_product_master(url):
    # Product Master มี 2 แถวหัวเรื่อง (Ownership, แผนก) อยู่เหนือแถวชื่อคอลัมน์จริง
    # แถวชื่อคอลัมน์จริงคือแถวที่ 3 (index 2)
    df = pd.read_csv(url, header=2)
    df.columns = [c.strip() for c in df.columns]
    required = ['SKUCode', 'PRODUCT']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"ไม่พบคอลัมน์: {', '.join(missing)} กรุณาตรวจสอบหัวตารางใน Google Sheets")
    df = df.dropna(subset=['SKUCode'])
    df['SKUCode'] = df['SKUCode'].astype(str).str.strip()
    return df


def find_header_row(rows, keys):
    for r in range(min(len(rows), 10)):
        row = rows[r]
        for c, v in enumerate(row):
            if isinstance(v, str) and v.strip().lower() in keys:
                return r, c
    return None, None


def _cell_blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ''


def _fill_merged_down(raw, ws, header_row):
    """ไฟล์บางแบบ (เช่น Route Planning สมุย) ใช้ merge cell แนวตั้ง: จุดส่ง 1 จุดมีหลาย PO
    แต่ Cust. ID / ชื่อลูกค้า / Ship-to ถูก merge ไว้แค่แถวบนสุด pandas จะเห็นแถวล่างๆ เป็นช่องว่าง
    -> แถว PO พวกนั้นจะไม่มี Cust ID และหลุดหายจากผลลัพธ์ ฟังก์ชันนี้เติมค่าจาก merge ลงไปทุกแถว
    ในช่วงนั้น (ทำเฉพาะ merge แนวตั้งที่อยู่ใต้หัวตาราง และเฉพาะเมื่อค่ามุมบนซ้ายตรงกันระหว่าง
    pandas กับ openpyxl — ถ้าไม่ตรงแปลว่าแถวไม่ตรงกัน ข้ามไปเลย ไม่เดา)"""
    for rng in ws.merged_cells.ranges:
        r0, c0 = rng.min_row - 1, rng.min_col - 1
        r1, c1 = rng.max_row - 1, rng.max_col - 1
        if r0 <= header_row or c0 != c1 or r1 >= len(raw):
            continue
        top = raw[r0][c0] if c0 < len(raw[r0]) else None
        anchor = ws.cell(row=rng.min_row, column=rng.min_col).value
        if _cell_blank(top) or anchor is None or str(top).strip() != str(anchor).strip():
            continue
        for r in range(r0 + 1, r1 + 1):
            if c0 < len(raw[r]) and _cell_blank(raw[r][c0]):
                raw[r][c0] = top


def process_route_file(file_bytes, master_df, selected_sheets=None, dedupe_shipto=True):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    # ใช้อ่าน merge cell (.xlsx เท่านั้น — .xls เปิดด้วย openpyxl ไม่ได้ ก็ข้ามไปตามปกติ)
    try:
        merge_wb = load_workbook(io.BytesIO(file_bytes))
    except Exception:
        merge_wb = None
    output_sheets = {}
    unique_stops_sheets = {}  # เวอร์ชัน "จุดส่งไม่ซ้ำ" คอลัมน์เหมือนไฟล์หลักทุกอย่าง แค่ยุบแถวพิกัดซ้ำ
    report = []
    unmatched = []
    ambiguous = []

    by_ship = master_df.drop_duplicates('norm_ship').set_index('norm_ship')
    by_code = master_df.drop_duplicates('norm_code').set_index('norm_code')

    ambiguous_ship_keys = set()
    if not dedupe_shipto:
        # หา Ship To Name ที่ซ้ำกันแต่พิกัด "ขัดแย้งกันจริง" (คนละสาขา) เพื่อไม่เดาสุ่มเลือกสาขาใดสาขาหนึ่ง
        grp = master_df.groupby('norm_ship')[['Latitude', 'Longitude']].nunique()
        ambiguous_ship_keys = set(grp[(grp['Latitude'] > 1) | (grp['Longitude'] > 1)].index)

    sheet_list = selected_sheets if selected_sheets is not None else xl.sheet_names

    for sheet_name in sheet_list:
        raw = xl.parse(sheet_name, header=None).values.tolist()
        header_row, cust_col = find_header_row(raw, [
            'cust code', 'cust id', 'cust. id', 'cust.id', 'cust. code', 'cust.code',
            'customer code', 'customers code', 'customer id', 'customers id',
        ])

        # ถ้าไม่เจอ Cust Code/ID เลย ลองหาคอลัมน์ทางเลือก (ไฟล์บางแบบ เช่น export สไตล์ invoice
        # ของโซนภูเก็ต/สมุย ไม่มีคอลัมน์ Cust Code เลย มีแต่ Cust. Name)
        if header_row is None:
            header_row, cust_col = find_header_row(raw, ['cust. name', 'cust name', 'cust.name', 'customer name', 'customers name'])

        if header_row is None:
            output_sheets[sheet_name] = xl.parse(sheet_name, header=None)
            report.append({'sheet': sheet_name, 'skipped': True, 'matched': 0, 'total': 0})
            continue

        if merge_wb is not None and sheet_name in merge_wb.sheetnames:
            _fill_merged_down(raw, merge_wb[sheet_name], header_row)

        header = raw[header_row]
        ship_col = None
        for i, v in enumerate(header):
            if isinstance(v, str) and 'ship' in v.lower() and 'name' in v.lower():
                ship_col = i
                break
        # ถ้าไม่เจอคอลัมน์ชื่อ "...ship...name..." ลองหา "Ship To" เฉยๆ (ไม่มีคำว่า name ต่อท้าย)
        # เผื่อไฟล์ตั้งชื่อคอลัมน์แบบนี้ (เจอในไฟล์ export บางโซน)
        if ship_col is None:
            for i, v in enumerate(header):
                if isinstance(v, str) and v.strip().lower() in ('ship to', 'ship-to', 'shipto'):
                    ship_col = i
                    break

        # คอลัมน์ Ship To มีอยู่ แต่ว่างทั้งชีท (เช่นชีท 23.09.26 ของไฟล์สมุย) = เท่ากับไม่มีคอลัมน์นี้
        # -> ตัดสินใจระดับชีทให้ไปใช้ชื่อลูกค้าแทน (ยังเป็นการจับด้วยชื่อ ไม่ใช่เดาด้วยรหัส)
        if ship_col is not None and all(
            ship_col >= len(rw) or _cell_blank(rw[ship_col]) for rw in raw[header_row + 1:]
        ):
            ship_col = None

        # ถ้าไม่มีคอลัมน์ ShipTo Name ในชีทนี้เลย ให้ใช้ Cust Name (หรือคอลัมน์ชื่ออื่นที่ใกล้เคียง) แทน
        name_fallback_col = None
        if ship_col is None:
            for i, v in enumerate(header):
                if isinstance(v, str) and 'cust' in v.lower() and 'name' in v.lower():
                    name_fallback_col = i
                    break
            if name_fallback_col is None:
                for i, v in enumerate(header):
                    if isinstance(v, str) and 'name' in v.lower() and 'ship' not in v.lower():
                        name_fallback_col = i
                        break

        # ระบุ "โหมด" การจับคู่ของทั้งชีทนี้ ครั้งเดียวตอนเจอ header (ไม่ใช่ทีละแถว):
        #   - ถ้ามีคอลัมน์ ShipTo Name หรือ Cust Name -> ใช้ชื่อจับคู่กับ Master ทุกแถวในชีทนี้
        #     (แถวไหน match ไม่ได้ ให้ถือว่า "ไม่มีพิกัด" ไปเลย ห้ามหลุดไปเดาด้วย Cust Code
        #      เพราะลูกค้า 1 รายอาจมีหลายสาขา การเดาด้วยโค้ดระดับลูกค้าจะได้พิกัดสาขาอื่นมาแบบผิดๆ)
        #   - ถ้าชีทนี้ไม่มีทั้งสองคอลัมน์เลย -> ทั้งชีทจะจับคู่ด้วย Cust Code <-> Ship To (Master) แทน
        use_code_fallback_for_sheet = (ship_col is None and name_fallback_col is None)

        out_rows = []
        matched = total = via_ship = via_code = 0

        for r, row in enumerate(raw):
            if r < header_row:
                out_rows.append(row)
                continue
            if r == header_row:
                new_header = list(row) + ['Latitude', 'Longitude']
                out_rows.append(new_header)
                continue

            cust_val = row[cust_col] if cust_col < len(row) else None
            if cust_val is None or (isinstance(cust_val, float) and pd.isna(cust_val)) or str(cust_val).strip() == '':
                continue

            code = str(cust_val).strip()
            total += 1
            if ship_col is not None and ship_col < len(row):
                ship_val = row[ship_col]
            elif name_fallback_col is not None and name_fallback_col < len(row):
                ship_val = row[name_fallback_col]
            else:
                ship_val = None

            hit = None
            is_ambiguous = False
            if not use_code_fallback_for_sheet:
                # โหมดจับคู่ด้วยชื่อ: ลองครั้งเดียว ไม่ fallback ไป Cust Code ต่อแถวต่อแถว
                if ship_val:
                    key = normalize(ship_val)
                    if key in ambiguous_ship_keys:
                        # ชื่อนี้ซ้ำในหลายสาขาที่พิกัดขัดแย้งกัน (เช่น Big C หลายสาขาชื่อเดียวกัน)
                        # ไม่เดาว่าเป็นสาขาไหน -> ถือว่ายังไม่มีพิกัด แล้วบันทึกแยกไว้ให้ตรวจสอบ
                        is_ambiguous = True
                    elif key in by_ship.index:
                        hit = by_ship.loc[key]
                        via_ship += 1
            else:
                # โหมดนี้ใช้เมื่อทั้งชีทไม่มีคอลัมน์ชื่อเลย -> จับคู่ด้วย Cust Code <-> Ship To (Master)
                key = code.upper()
                if key in by_code.index:
                    hit = by_code.loc[key]
                    via_code += 1

            new_row = list(row)
            if hit is not None:
                matched += 1
                lat = hit['Latitude'].iloc[0] if hasattr(hit['Latitude'], 'iloc') else hit['Latitude']
                lon = hit['Longitude'].iloc[0] if hasattr(hit['Longitude'], 'iloc') else hit['Longitude']
                new_row += [lat, lon]
            elif is_ambiguous:
                new_row += [None, None]
                dup_rows = master_df[master_df['norm_ship'] == normalize(ship_val)]
                branch_count = len(dup_rows.drop_duplicates(subset=['Latitude', 'Longitude']))
                ambiguous.append({
                    'Cust Code': code,
                    'Ship To Name': ship_val or '',
                    'Sheet': sheet_name,
                    'จำนวนสาขาที่ชื่อซ้ำ': branch_count,
                    'พบเมื่อ': datetime.now().strftime('%Y-%m-%d'),
                })
            else:
                new_row += [None, None]
                unmatched.append({
                    'Cust Code': code,
                    'Ship To Name': ship_val or '',
                    'Sheet': sheet_name,
                    'พบเมื่อ': datetime.now().strftime('%Y-%m-%d'),

                })
            out_rows.append(new_row)

        output_sheets[sheet_name] = pd.DataFrame(out_rows)
        report.append({'sheet': sheet_name, 'skipped': False, 'matched': matched, 'total': total,
                        'via_ship': via_ship, 'via_code': via_code})

        # ---- สร้างเวอร์ชัน "จุดส่งไม่ซ้ำ" ของชีทนี้ ----
        # ใช้ template เดียวกับ out_rows เป๊ะ (คอลัมน์เดิมทั้งหมด + Latitude/Longitude)
        # แค่ยุบแถวที่พิกัด (2 คอลัมน์สุดท้าย) ซ้ำกันให้เหลือแถวแรกที่เจอ แล้วเติมคอลัมน์นับจำนวนออเดอร์ต่อท้าย
        preamble_and_header = out_rows[:header_row + 1]
        data_rows = out_rows[header_row + 1:]

        seen_order = []
        seen_row = {}
        seen_count = {}
        no_coord_counter = 0
        for row in data_rows:
            lat_val, lon_val = row[-2], row[-1]
            if lat_val is None or lon_val is None:
                # ไม่มีพิกัด -> ยังคงเก็บแถวนี้ไว้เหมือนเดิม (ห้ามลบทิ้งเงียบๆ) แค่ไม่เอาไปยุบรวมกับแถวไหน
                # เพราะไม่รู้ว่ามันคือจุดเดียวกับแถวอื่นไหนหรือเปล่า
                key = ('__NOCOORD__', no_coord_counter)
                no_coord_counter += 1
            else:
                key = (lat_val, lon_val)
            if key not in seen_row:
                seen_row[key] = row
                seen_count[key] = 1
                seen_order.append(key)
            else:
                seen_count[key] += 1

        unique_header = list(preamble_and_header[-1]) + ['จำนวนออเดอร์รวม'] if preamble_and_header else \
            list(raw[header_row]) + ['Latitude', 'Longitude', 'จำนวนออเดอร์รวม']
        unique_data_rows = [list(seen_row[k]) + [seen_count[k]] for k in seen_order]
        unique_sheet_rows = preamble_and_header[:-1] + [unique_header] + unique_data_rows if preamble_and_header else \
            [unique_header] + unique_data_rows
        unique_stops_sheets[sheet_name] = pd.DataFrame(unique_sheet_rows)
        report[-1]['unique_stops'] = len(unique_data_rows)

    return output_sheets, report, unmatched, ambiguous, unique_stops_sheets


def process_weight_file(file_bytes, product_df, selected_sheets=None):
    """เชื่อม Product Code (ไฟล์ route) กับ SKUCode (Product Master) เพื่อคำนวณน้ำหนักทีละแถว
    เก็บคอลัมน์เดิมทั้งหมดไว้ครบ (เหมือนโหมดพิกัด) แค่เพิ่ม 'Weight (kg)' ต่อท้าย"""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    output_sheets = {}
    report = []
    no_weight = []

    weight_lookup = product_df.dropna(subset=['SKUCode']).drop_duplicates('SKUCode').set_index('SKUCode')

    sheet_list = selected_sheets if selected_sheets is not None else xl.sheet_names

    for sheet_name in sheet_list:
        raw = xl.parse(sheet_name, header=None).values.tolist()
        header_row, product_col = find_header_row(raw, ['product code'])

        if header_row is None:
            output_sheets[sheet_name] = xl.parse(sheet_name, header=None)
            report.append({'sheet': sheet_name, 'skipped': True, 'calculated': 0, 'total': 0})
            continue

        header = raw[header_row]
        qty_col = None
        units_col = None
        for i, v in enumerate(header):
            if isinstance(v, str) and v.strip().lower() == 'qty':
                qty_col = i
            if isinstance(v, str) and v.strip().lower() == 'units':
                units_col = i

        out_rows = []
        calculated = total = 0

        for r, row in enumerate(raw):
            if r < header_row:
                out_rows.append(row)
                continue
            if r == header_row:
                out_rows.append(list(row) + ['Weight (kg)'])
                continue

            product_val = row[product_col] if product_col < len(row) else None
            if product_val is None or (isinstance(product_val, float) and pd.isna(product_val)) or str(product_val).strip() == '':
                # ไม่มี Product Code -> ยังคงเก็บแถวนี้ไว้เหมือนเดิม (ห้ามลบทิ้งเงียบๆ) แค่ไม่มีน้ำหนักให้คำนวณ
                out_rows.append(list(row) + [None])
                continue

            code = str(product_val).strip()
            total += 1
            qty_val = row[qty_col] if (qty_col is not None and qty_col < len(row)) else None
            units_val = row[units_col] if (units_col is not None and units_col < len(row)) else None
            units_str = str(units_val).strip().upper() if units_val is not None and not (isinstance(units_val, float) and pd.isna(units_val)) else ''

            weight = None
            reason = None

            if code not in weight_lookup.index:
                reason = 'ไม่พบ Product Code นี้ใน Product Master'
            elif qty_val is None or (isinstance(qty_val, float) and pd.isna(qty_val)):
                reason = 'ไม่มีค่า Qty ในแถวนี้'
            elif units_str == 'PCS':
                unit_wt = weight_lookup.loc[code, 'Unit Net Wt (kg)']
                if hasattr(unit_wt, 'iloc'):
                    unit_wt = unit_wt.iloc[0]
                if pd.isna(unit_wt):
                    reason = 'มี Product Code แต่ Product Master ยังไม่มีข้อมูลน้ำหนัก (Unit Net Wt)'
                else:
                    weight = float(qty_val) * float(unit_wt)
            elif units_str == 'KG':
                weight = float(qty_val)  # Qty เป็นน้ำหนักอยู่แล้ว ไม่ต้องคูณ
            else:
                reason = f'หน่วย (Units) "{units_val}" ยังไม่รองรับการคำนวณอัตโนมัติ'

            new_row = list(row) + [weight]
            if weight is not None:
                calculated += 1
            else:
                no_weight.append({
                    'Product Code': code,
                    'Qty': qty_val,
                    'Units': units_val if units_val is not None else '',
                    'Sheet': sheet_name,
                    'สาเหตุ': reason,
                })
            out_rows.append(new_row)

        output_sheets[sheet_name] = pd.DataFrame(out_rows)
        report.append({'sheet': sheet_name, 'skipped': False, 'calculated': calculated, 'total': total})

    return output_sheets, report, no_weight


def _is_blank(v):
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == '':
        return True
    return False


def _drop_empty_columns(df):
    """ลบคอลัมน์ที่ว่างเปล่าทุกแถว (ทั้ง header และข้อมูล) เช่นคอลัมน์ spacer ระหว่างคอลัมน์จริงในไฟล์ route ต้นฉบับ"""
    keep_cols = [c for c in df.columns if not df[c].apply(_is_blank).all()]
    if not keep_cols:  # กันเคส edge case ที่ทุกคอลัมน์ว่างหมด (ไม่ควรเกิดขึ้นจริง)
        return df
    return df[keep_cols].reset_index(drop=True)


def to_excel_bytes(sheets_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in sheets_dict.items():
            df_clean = _drop_empty_columns(df)
            df_clean.to_excel(writer, sheet_name=name[:31], header=False, index=False)
            ws = writer.sheets[name[:31]]
            # ปรับความกว้างคอลัมน์ให้พอดีกับตัวอักษรที่ยาวที่สุดในคอลัมน์นั้น (openpyxl ไม่มี autofit ในตัว จำลองเอง)
            for col_idx in range(df_clean.shape[1]):
                max_len = 0
                for val in df_clean.iloc[:, col_idx]:
                    if not _is_blank(val):
                        max_len = max(max_len, len(str(val)))
                col_letter = get_column_letter(col_idx + 1)
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 60)
    return output.getvalue()


def to_csv_bytes(sheets_dict):
    # รวมทุก sheet ที่เลือกเป็น CSV เดียว คั่นด้วยชื่อ sheet
    output = io.StringIO()
    for name, df in sheets_dict.items():
        output.write(f"# Sheet: {name}\n")
        df.to_csv(output, header=False, index=False)
        output.write("\n")
    return output.getvalue().encode('utf-8-sig')


def to_kml_bytes(sheets_dict):
    # หาแถว header จริง (แถวที่มีคำว่า Latitude/Longitude) เพราะบาง sheet มีแถวหัวเรื่องอยู่ด้านบน header จริง
    placemarks = []
    for sheet_name, df in sheets_dict.items():
        if df.empty or df.shape[1] < 2:
            continue

        header_row_idx = None
        lat_idx = lon_idx = None
        for i in range(min(len(df), 15)):
            row_vals = df.iloc[i].tolist()
            if 'Latitude' in row_vals and 'Longitude' in row_vals:
                header_row_idx = i
                lat_idx = row_vals.index('Latitude')
                lon_idx = row_vals.index('Longitude')
                break
        if header_row_idx is None:
            continue

        header = df.iloc[header_row_idx].tolist()
        name_idx = 0
        for i, h in enumerate(header):
            if isinstance(h, str) and 'name' in h.lower():
                name_idx = i
                break

        for _, row in df.iloc[header_row_idx + 1:].iterrows():
            lat, lon = row.iloc[lat_idx], row.iloc[lon_idx]
            if pd.isna(lat) or pd.isna(lon):
                continue
            label = str(row.iloc[name_idx]) if name_idx < len(row) else "Stop"
            placemarks.append(f"""
    <Placemark>
      <name>{label}</name>
      <description>Sheet: {sheet_name}</description>
      <Point><coordinates>{lon},{lat},0</coordinates></Point>
    </Placemark>""")

    kml = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Route Coordinates</name>
    {''.join(placemarks)}
  </Document>
</kml>"""
    return kml.encode('utf-8')


def unmatched_to_excel_bytes(unmatched_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        unmatched_df.to_excel(writer, sheet_name='Unmatched', index=False)
        ws = writer.sheets['Unmatched']
        for i, col in enumerate(unmatched_df.columns, start=1):
            ws.column_dimensions[chr(64 + i)].width = 28
    return output.getvalue()


def stat_box(label, value, sub=""):
    st.markdown(f"""
    <div class="rcm-stat">
        <div class="rcm-stat-label">{label}</div>
        <div class="rcm-stat-value">{value}</div>
        <div class="rcm-stat-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


# ============================================================
# HEADER
# ============================================================
st.markdown("""
<div class="rcm-header">
    <svg class="rcm-route-svg" viewBox="0 0 400 220" xmlns="http://www.w3.org/2000/svg">
        <path d="M 60 190 Q 150 170 180 120 T 300 40" stroke="rgba(255,255,255,0.22)" stroke-width="1.5" fill="none" stroke-dasharray="4 5"/>
        <path d="M 90 200 Q 170 150 220 130 T 330 55" stroke="rgba(255,255,255,0.12)" stroke-width="1" fill="none"/>
        <circle cx="300" cy="40" r="7" fill="#FF6B35" opacity="0.95">
            <animate attributeName="opacity" values="0.6;1;0.6" dur="2.4s" repeatCount="indefinite"/>
        </circle>
        <circle cx="300" cy="40" r="13" fill="none" stroke="#FF6B35" stroke-width="1.5" opacity="0.5"/>
        <circle cx="60" cy="190" r="6" fill="#2F80ED" opacity="0.95">
            <animate attributeName="opacity" values="0.6;1;0.6" dur="2.4s" repeatCount="indefinite" begin="1.2s"/>
        </circle>
        <circle cx="60" cy="190" r="11" fill="none" stroke="#2F80ED" stroke-width="1.5" opacity="0.5"/>
        <circle cx="180" cy="120" r="3.5" fill="#FFB347" opacity="0.8"/>
    </svg>
    <div class="rcm-header-content">
        <div class="rcm-eyebrow">🚚 FLEET OPERATIONS · GEO-MATCHING SYSTEM</div>
        <div class="rcm-title">Route Coordinate Matcher</div>
        <div class="rcm-subtitle">อัปโหลดไฟล์ route เพื่อจับคู่พิกัดกับ Master Data</div>
        <div class="rcm-features">
            <div class="rcm-feature">
                <div class="rcm-feature-icon">🎯</div>
                <div class="rcm-feature-text">
                    <div class="rcm-feature-title">Geo-Matching</div>
                    <div class="rcm-feature-sub">จับคู่พิกัดอัตโนมัติ</div>
                </div>
            </div>
            <div class="rcm-feature">
                <div class="rcm-feature-icon">🗄️</div>
                <div class="rcm-feature-text">
                    <div class="rcm-feature-title">Master Data</div>
                    <div class="rcm-feature-sub">ฐานข้อมูลลูกค้า</div>
                </div>
            </div>
            <div class="rcm-feature">
                <div class="rcm-feature-icon">🚛</div>
                <div class="rcm-feature-text">
                    <div class="rcm-feature-title">Fleet Efficiency</div>
                    <div class="rcm-feature-sub">เพิ่มประสิทธิภาพการขนส่ง</div>
                </div>
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

tab_process, tab_product, tab_qr, tab_bkk = st.tabs(["📤  ประมวลผลไฟล์", "📦  Product Master", "🔗  QR Code", "🏙️  For BKK"])

with tab_process:
    if not MASTER_SHEET_CSV_URL:
        st.error("⚠️ ยังไม่ได้ตั้งค่า MASTER_SHEET_CSV_URL — ดูวิธีตั้งค่าใน DEPLOY_INSTRUCTIONS.md")
        st.stop()

    try:
        master_df = load_master_data(MASTER_SHEET_CSV_URL)
        unique_customer_count = master_df['norm_ship'].nunique()
        status_col, sync_col = st.columns([5, 1])
        with status_col:
            st.markdown(f"""
            <div class="rcm-status-card">
                <div class="rcm-status-left">
                    <div class="rcm-status-icon">✅</div>
                    <div>
                        <div class="rcm-status-title">เชื่อมต่อฐานข้อมูลสำเร็จ</div>
                        <div class="rcm-status-sub">{unique_customer_count:,} รายชื่อลูกค้าในระบบ (ไม่ซ้ำ)</div>
                    </div>
                </div>
                <div class="rcm-status-right">
                    <div class="rcm-status-right-icon">🗃️</div>
                    <div class="rcm-status-pill"><span class="rcm-status-dot"></span>Connected</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with sync_col:
            st.markdown('<div class="rcm-sync-btn-wrap">', unsafe_allow_html=True)
            if st.button("🔄 ซิงค์ตอนนี้", use_container_width=True, help="ดึงข้อมูล Master Data ล่าสุดจาก Google Sheets ทันที"):
                load_master_data.clear()
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
    except Exception as e:
        st.error(f"โหลด Master Data ไม่สำเร็จ: {e}")
        st.stop()

    if 'show_master_preview' not in st.session_state:
        st.session_state['show_master_preview'] = False

    chevron = "▾" if st.session_state['show_master_preview'] else "▸"
    st.markdown('<div class="rcm-preview-toggle-wrap">', unsafe_allow_html=True)
    preview_clicked = st.button(
        f"📄   ดูตัวอย่าง Master Data — ตรวจสอบข้อมูลลูกค้าและเส้นทางก่อนประมวลผล   {chevron}",
        key="master_preview_toggle",
        use_container_width=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)
    if preview_clicked:
        st.session_state['show_master_preview'] = not st.session_state['show_master_preview']
        st.rerun()

    if st.session_state['show_master_preview']:
        st.dataframe(master_df[['Ship To', 'Ship To Name', 'Latitude', 'Longitude']].head(20), use_container_width=True)

    st.write("")

    uploaded_file = st.file_uploader("อัพโหลดไฟล์ route", type=['xls', 'xlsx'], label_visibility="collapsed")

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()

        try:
            xl_peek = pd.ExcelFile(io.BytesIO(file_bytes))
            all_sheet_names = xl_peek.sheet_names
        except Exception as e:
            st.error(f"เปิดไฟล์ไม่ได้: {e}")
            st.stop()

        st.markdown("##### 📑 เลือก Sheet ที่ต้องการประมวลผล")
        sheet_cols = st.columns(min(len(all_sheet_names), 4))
        selected_sheets = []
        for i, sn in enumerate(all_sheet_names):
            with sheet_cols[i % len(sheet_cols)]:
                checked = st.checkbox(sn, value=True, key=f"sheet_{sn}")
                if checked:
                    selected_sheets.append(sn)

        if not selected_sheets:
            st.warning("⚠️ กรุณาเลือกอย่างน้อย 1 Sheet เพื่อเริ่มประมวลผล")
            st.stop()

        dedupe_shipto = st.checkbox(
            "✨ ทำให้ Ship To Name ไม่ซ้ำกันอัตโนมัติ (เลือกแถวแรกที่เจอ)",
            value=True,
            help="เปิด (ค่าเริ่มต้น): ถ้าชื่อซ้ำกันในระบบ จะเลือกแถวแรกที่เจอมาใช้ทันที เร็ว แต่ถ้าเป็นเชนร้านที่มีหลายสาขาชื่อเดียวกัน (เช่น Big C) อาจได้พิกัดผิดสาขา\n\nปิด: ถ้าเจอชื่อซ้ำที่พิกัดขัดแย้งกันจริง (คนละสาขา) จะไม่เดา — ถือว่ายังไม่มีพิกัด แล้วแยกไปอยู่ลิสต์ 'ชื่อซ้ำ-ต้องตรวจสอบ' แทน ปลอดภัยกว่าแต่ match ได้น้อยลง",
        )

        st.write("")
        process_clicked = st.button("🚀 ประมวลผล Sheet ที่เลือก", type="primary", use_container_width=True)

        if process_clicked:
            with st.spinner("🛰️ กำลังจับคู่พิกัด..."):
                try:
                    sheets, report, unmatched, ambiguous, unique_stops_sheets = process_route_file(
                        file_bytes, master_df, selected_sheets, dedupe_shipto=dedupe_shipto
                    )
                    st.session_state['last_sheets'] = sheets
                    st.session_state['last_report'] = report
                    st.session_state['last_unmatched'] = unmatched
                    st.session_state['last_ambiguous'] = ambiguous
                    st.session_state['last_unique_stops_sheets'] = unique_stops_sheets
                    st.session_state['last_filename'] = uploaded_file.name
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")
                    st.exception(e)

        if 'last_sheets' in st.session_state and st.session_state.get('last_filename') == uploaded_file.name:
            sheets = st.session_state['last_sheets']
            report = st.session_state['last_report']
            unmatched = st.session_state['last_unmatched']
            ambiguous = st.session_state.get('last_ambiguous', [])
            unique_stops_sheets = st.session_state.get('last_unique_stops_sheets', {})

            total_matched = sum(r['matched'] for r in report)
            total_rows = sum(r['total'] for r in report)
            match_rate = (total_matched / total_rows * 100) if total_rows else 0

            card_class = "rcm-card-success" if match_rate >= 90 else "rcm-card-warn"
            st.markdown(f"""
            <div class="rcm-card {card_class}">
                🎯 <b>เสร็จแล้ว!</b> จับคู่พิกัดได้ {total_matched:,} / {total_rows:,} แถว ({match_rate:.1f}%)
            </div>
            """, unsafe_allow_html=True)

            cols = st.columns(len(report))
            for col, r in zip(cols, report):
                with col:
                    if r['skipped']:
                        stat_box(f"SHEET: {r['sheet']}", "ข้าม", "ไม่พบคอลัมน์ Cust Code")
                    else:
                        stat_box(f"SHEET: {r['sheet']}", f"{r['matched']}/{r['total']}",
                                  f"SHIP-TO {r['via_ship']} · CODE {r['via_code']}")

            st.markdown('<div class="rcm-glow-divider"></div>', unsafe_allow_html=True)

            # --- Export format selection ---
            st.markdown("##### 💾 เลือกฟอร์แมตไฟล์ผลลัพธ์")
            export_format = st.radio(
                "รูปแบบไฟล์",
                options=["Excel (.xlsx)", "CSV (.csv)", "KML (.kml — เปิดใน Google Earth/Maps)"],
                horizontal=True,
                label_visibility="collapsed",
            )

            base_name = uploaded_file.name.rsplit('.', 1)[0]
            dl_col1, dl_col2, dl_col3 = st.columns(3)

            with dl_col1:
                if export_format.startswith("Excel"):
                    data_bytes = to_excel_bytes(sheets)
                    fname = f"{base_name}_with_coordinates.xlsx"
                    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                elif export_format.startswith("CSV"):
                    data_bytes = to_csv_bytes(sheets)
                    fname = f"{base_name}_with_coordinates.csv"
                    mime = "text/csv"
                else:
                    data_bytes = to_kml_bytes(sheets)
                    fname = f"{base_name}_with_coordinates.kml"
                    mime = "application/vnd.google-earth.kml+xml"

                st.download_button(
                    "⬇️ ดาวน์โหลดไฟล์พร้อมพิกัด",
                    data=data_bytes,
                    file_name=fname,
                    mime=mime,
                    use_container_width=True,
                )

            unmatched_df = pd.DataFrame(unmatched).drop_duplicates(subset=['Cust Code']) if unmatched else pd.DataFrame()

            with dl_col2:
                if not unmatched_df.empty:
                    unmatched_bytes = unmatched_to_excel_bytes(unmatched_df)
                    st.download_button(
                        f"📋 ดาวน์โหลดรายการที่ยังไม่มีพิกัด ({len(unmatched_df)})",
                        data=unmatched_bytes,
                        file_name=f"unmatched_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                else:
                    st.markdown("""
                    <div class="rcm-card rcm-card-success" style="text-align:center;">
                        ✅ ไม่มีรายการตกหล่น — พิกัดครบทุกแถว
                    </div>
                    """, unsafe_allow_html=True)

            with dl_col3:
                total_unique_stops = sum(r.get('unique_stops', 0) for r in report)
                if unique_stops_sheets and total_unique_stops > 0:
                    stops_bytes = to_excel_bytes(unique_stops_sheets)
                    st.download_button(
                        f"📍 จุดส่งไม่ซ้ำ สำหรับ OptimoRoute ({total_unique_stops})",
                        data=stops_bytes,
                        file_name=f"{base_name}_unique_stops.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        help="คอลัมน์เหมือนไฟล์หลักทุกอย่าง เพิ่มแค่คอลัมน์ 'จำนวนออเดอร์รวม' — ยุบแถวที่พิกัดซ้ำกันให้เหลือ 1 แถวต่อ 1 จุด",
                    )

            if not unmatched_df.empty:
                st.write("")
                st.markdown(f"""
                <div class="rcm-card rcm-card-warn">
                    ⚠️ <b>มี {len(unmatched_df)} รายการที่ยังไม่มีพิกัด</b> — ใช้ไฟล์ที่ดาวน์โหลดด้านบนไปเพิ่มลงใน Master Data ได้เลย (ที่ Google Sheets โดยตรง)
                </div>
                """, unsafe_allow_html=True)
                st.dataframe(unmatched_df, use_container_width=True, hide_index=True)

            ambiguous_df = pd.DataFrame(ambiguous).drop_duplicates(subset=['Cust Code']) if ambiguous else pd.DataFrame()
            if not ambiguous_df.empty:
                st.write("")
                st.markdown(f"""
                <div class="rcm-card rcm-card-warn">
                    🔀 <b>พบ {len(ambiguous_df)} รายการที่ชื่อซ้ำในหลายสาขา (พิกัดขัดแย้งกัน)</b> —
                    ระบบไม่กล้าเดาว่าเป็นสาขาไหน กรุณาเปิด Master Data แล้วแก้ Ship To Name ให้ระบุสาขาชัดเจน (เช่น เพิ่มชื่อสาขาต่อท้าย) แล้วลองประมวลผลใหม่
                </div>
                """, unsafe_allow_html=True)
                st.dataframe(ambiguous_df, use_container_width=True, hide_index=True)
                ambiguous_bytes = unmatched_to_excel_bytes(ambiguous_df)
                st.download_button(
                    f"🔀 ดาวน์โหลดรายการชื่อซ้ำที่ต้องตรวจสอบ ({len(ambiguous_df)})",
                    data=ambiguous_bytes,
                    file_name=f"ambiguous_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    st.markdown('<div class="rcm-glow-divider"></div>', unsafe_allow_html=True)
    st.caption(f"🛰️ Master Data sync (cache): {datetime.now().strftime('%Y-%m-%d %H:%M')} — รีเฟรชอัตโนมัติทุก 5 นาที หรือกด \"ซิงค์ตอนนี้\" ด้านบนเพื่อดึงข้อมูลล่าสุดทันที")

with tab_product:
    if not PRODUCT_SHEET_CSV_URL:
        st.markdown("""
        <div class="rcm-card rcm-card-warn">
            ⚠️ <b>ยังไม่ได้ตั้งค่า PRODUCT_SHEET_CSV_URL</b><br>
            เพิ่มใน Streamlit Secrets: <code>PRODUCT_SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/YOUR_ID/export?format=csv&gid=0"</code>
        </div>
        """, unsafe_allow_html=True)
    else:
        try:
            product_df = load_product_master(PRODUCT_SHEET_CSV_URL)
            pstatus_col, psync_col = st.columns([5, 1])
            with pstatus_col:
                st.markdown(f"""
                <div class="rcm-status-card">
                    <div class="rcm-status-left">
                        <div class="rcm-status-icon">✅</div>
                        <div>
                            <div class="rcm-status-title">เชื่อมต่อ Product Master สำเร็จ</div>
                            <div class="rcm-status-sub">{len(product_df):,} รายการสินค้าในระบบ</div>
                        </div>
                    </div>
                    <div class="rcm-status-right">
                        <div class="rcm-status-right-icon">📦</div>
                        <div class="rcm-status-pill"><span class="rcm-status-dot"></span>Connected</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with psync_col:
                st.markdown('<div class="rcm-sync-btn-wrap">', unsafe_allow_html=True)
                if st.button("🔄 ซิงค์ตอนนี้", use_container_width=True, key="product_sync_btn",
                             help="ดึงข้อมูล Product Master ล่าสุดจาก Google Sheets ทันที"):
                    load_product_master.clear()
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"โหลด Product Master ไม่สำเร็จ: {e}")
            st.stop()

        st.write("")
        search_term = st.text_input(
            "ค้นหาสินค้า", label_visibility="collapsed",
            placeholder="🔍 ค้นหาด้วย SKUCode หรือชื่อสินค้า...",
        )

        display_cols = ['SKUCode', 'PRODUCT', 'BRAND', 'CATEGORY', 'UOM',
                         'Pieces per Cases/Carton', 'Unit Net Wt (kg)',
                         'Case Weight (kg)', 'Case Volume (CBM)', 'Status']
        display_cols = [c for c in display_cols if c in product_df.columns]
        display_df = product_df[display_cols]

        if search_term:
            mask = (
                display_df['SKUCode'].astype(str).str.contains(search_term, case=False, na=False) |
                display_df['PRODUCT'].astype(str).str.contains(search_term, case=False, na=False)
            )
            display_df = display_df[mask]

        st.dataframe(display_df, use_container_width=True, hide_index=True, height=460)
        st.caption(f"แสดง {len(display_df):,} จาก {len(product_df):,} รายการทั้งหมด")

        st.markdown('<div class="rcm-glow-divider"></div>', unsafe_allow_html=True)
        st.markdown("##### ⚖️ คำนวณน้ำหนักจากไฟล์ route")
        st.caption("โยนไฟล์ route เข้ามา ระบบจะเชื่อม Product Code กับ Product Master แล้วเติมคอลัมน์น้ำหนักท้ายทุกแถว (คอลัมน์เดิมครบเหมือนไฟล์ต้นฉบับ)")

        weight_file = st.file_uploader(
            "อัพโหลดไฟล์ route สำหรับคำนวณน้ำหนัก", type=['xls', 'xlsx'],
            label_visibility="collapsed", key="weight_file_uploader",
        )

        if weight_file is not None:
            wfile_bytes = weight_file.read()
            try:
                wxl_peek = pd.ExcelFile(io.BytesIO(wfile_bytes))
                w_sheet_names = wxl_peek.sheet_names
            except Exception as e:
                st.error(f"เปิดไฟล์ไม่ได้: {e}")
                st.stop()

            st.write("")
            wsheet_cols = st.columns(min(len(w_sheet_names), 4))
            w_selected_sheets = []
            for i, sn in enumerate(w_sheet_names):
                with wsheet_cols[i % len(wsheet_cols)]:
                    checked = st.checkbox(sn, value=True, key=f"wsheet_{sn}")
                    if checked:
                        w_selected_sheets.append(sn)

            if not w_selected_sheets:
                st.warning("⚠️ กรุณาเลือกอย่างน้อย 1 Sheet เพื่อเริ่มประมวลผล")
                st.stop()

            st.write("")
            weight_process_clicked = st.button("⚖️ คำนวณน้ำหนัก", type="primary", use_container_width=True, key="weight_process_btn")

            if weight_process_clicked:
                with st.spinner("⚖️ กำลังคำนวณน้ำหนัก..."):
                    try:
                        w_sheets, w_report, w_no_weight = process_weight_file(wfile_bytes, product_df, w_selected_sheets)
                        st.session_state['last_w_sheets'] = w_sheets
                        st.session_state['last_w_report'] = w_report
                        st.session_state['last_w_no_weight'] = w_no_weight
                        st.session_state['last_w_filename'] = weight_file.name
                    except Exception as e:
                        st.error(f"เกิดข้อผิดพลาด: {e}")
                        st.exception(e)

            if 'last_w_sheets' in st.session_state and st.session_state.get('last_w_filename') == weight_file.name:
                w_sheets = st.session_state['last_w_sheets']
                w_report = st.session_state['last_w_report']
                w_no_weight = st.session_state['last_w_no_weight']

                w_total_calc = sum(r['calculated'] for r in w_report)
                w_total_rows = sum(r['total'] for r in w_report)
                w_rate = (w_total_calc / w_total_rows * 100) if w_total_rows else 0

                w_card_class = "rcm-card-success" if w_rate >= 90 else "rcm-card-warn"
                st.markdown(f"""
                <div class="rcm-card {w_card_class}">
                    ⚖️ <b>เสร็จแล้ว!</b> คำนวณน้ำหนักได้ {w_total_calc:,} / {w_total_rows:,} แถว ({w_rate:.1f}%)
                </div>
                """, unsafe_allow_html=True)

                w_cols = st.columns(len(w_report))
                for col, r in zip(w_cols, w_report):
                    with col:
                        if r['skipped']:
                            stat_box(f"SHEET: {r['sheet']}", "ข้าม", "ไม่พบคอลัมน์ Product Code")
                        else:
                            stat_box(f"SHEET: {r['sheet']}", f"{r['calculated']}/{r['total']}", "คำนวณสำเร็จ")

                st.write("")
                w_dl_col1, w_dl_col2 = st.columns(2)
                w_base_name = weight_file.name.rsplit('.', 1)[0]

                with w_dl_col1:
                    w_output_bytes = to_excel_bytes(w_sheets)
                    st.download_button(
                        "⬇️ ดาวน์โหลดไฟล์พร้อมน้ำหนัก",
                        data=w_output_bytes,
                        file_name=f"{w_base_name}_with_weight.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="weight_download_main",
                    )

                w_no_weight_df = pd.DataFrame(w_no_weight).drop_duplicates(subset=['Product Code', 'สาเหตุ']) if w_no_weight else pd.DataFrame()

                with w_dl_col2:
                    if not w_no_weight_df.empty:
                        w_nw_bytes = unmatched_to_excel_bytes(w_no_weight_df)
                        st.download_button(
                            f"📋 รายการที่ยังไม่มีน้ำหนัก ({len(w_no_weight_df)})",
                            data=w_nw_bytes,
                            file_name=f"no_weight_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="weight_download_missing",
                        )
                    else:
                        st.markdown("""
                        <div class="rcm-card rcm-card-success" style="text-align:center;">
                            ✅ คำนวณน้ำหนักได้ครบทุกแถว
                        </div>
                        """, unsafe_allow_html=True)

                if not w_no_weight_df.empty:
                    st.write("")
                    st.markdown(f"""
                    <div class="rcm-card rcm-card-warn">
                        ⚠️ <b>มี {len(w_no_weight_df)} รายการที่ยังคำนวณน้ำหนักไม่ได้</b> — ดูคอลัมน์ "สาเหตุ" เพื่อรู้ว่าต้องแก้อะไร (ส่วนใหญ่คือ Product Master ยังไม่มีข้อมูลน้ำหนักของ SKU นั้น)
                    </div>
                    """, unsafe_allow_html=True)
                    st.dataframe(w_no_weight_df, use_container_width=True, hide_index=True)


# ============================================================
# TAB 3: QR CODE GENERATOR
# ============================================================
def generate_qr_bytes(link, box_size=10, border=4, fill_color='#000000', back_color='#FFFFFF'):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=box_size,
        border=border,
    )
    qr.add_data(link)
    qr.make(fit=True)
    img = qr.make_image(fill_color=fill_color, back_color=back_color).convert('RGB')
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return buf.getvalue(), img.size


# ---- Batch mode: read links straight out of an uploaded file, embed a QR per row ----
QR_URL_RE = re.compile(r'https?://\S+')


def find_link_column(sheet_rows):
    """sheet_rows: list of row tuples (raw cell values, 0-indexed).
    Scans every cell (content, not header text) and returns the 0-based column
    index with the most URL-looking values — or None if the sheet has none.
    Content-based on purpose: header names for a maps-link column vary a lot
    between files ('Ship to Google Plus code (Epicor)', 'ที่ตั้ง', a column with
    no header at all, ...), so we trust what's actually in the cells."""
    counts = {}
    for row in sheet_rows:
        for i, v in enumerate(row):
            if isinstance(v, str) and QR_URL_RE.search(v):
                counts[i] = counts.get(i, 0) + 1
    if not counts:
        return None
    return max(counts, key=counts.get)


def _guess_table_top(raw_rows):
    """หาแถวหัวตาราง (0-indexed): แถวแรกที่มีข้อมูลเกินครึ่งของแถวที่กว้างที่สุด (อย่างน้อย 3 ช่อง)
    ใช้แค่กำหนดตำแหน่งวาง QR — ถ้าหาไม่เจอคืน None"""
    counts = [sum(1 for v in row if not (v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ''))
              for row in raw_rows]
    if not counts:
        return None
    need = max(3, max(counts) / 2)
    for i, c in enumerate(counts):
        if c >= need:
            return i
    return None


def _embed_qr_column(ws, raw_rows, link_col, start_col0, box_size=5, border=2):
    """Appends a 'QR Code' column at 0-based column index start_col0, embedding
    a QR image at NATIVE resolution (no shrinking) for every row whose link_col
    cell contains a link — long multi-stop Google Maps links need a fairly
    dense/large QR to stay scannable, so images are never downscaled to a fixed
    thumbnail. Never touches any other existing cell. Returns qr_count."""
    new_col_letter = get_column_letter(start_col0 + 1)
    ws.cell(row=1, column=start_col0 + 1, value="QR Code")

    links = []
    for r_idx, row in enumerate(raw_rows):
        v = row[link_col] if link_col < len(row) else None
        if isinstance(v, str):
            m = QR_URL_RE.search(v)
            if m:
                links.append((r_idx, m.group(0)))

    # ลิงก์เดียวทั้งชีท (เช่นลิงก์เส้นทางท้ายชีทของไฟล์ OptimoRoute) -> วาง QR ลอยไว้ขวาบนข้างตาราง
    # เริ่มที่แถวข้อมูลแรก และไม่ยืดความสูงแถวไหนเลย (แถวลิงก์ด้านล่างสูงปกติ)
    # ลิงก์หลายแถว (เช่น Master Data) -> วาง QR ข้างแถวของมันเหมือนเดิม จะได้รู้ว่า QR ไหนของแถวไหน
    if len(links) == 1:
        r_idx, url = links[0]
        try:
            qr_bytes, (w_px, h_px) = generate_qr_bytes(url, box_size=box_size, border=border)
        except Exception:
            return 0
        top = _guess_table_top(raw_rows)
        anchor_row = min(top + 2, r_idx + 1) if top is not None else 2  # แถวข้อมูลแรก (1-indexed)
        ws.add_image(XLImage(io.BytesIO(qr_bytes)), f"{new_col_letter}{anchor_row}")
        ws.column_dimensions[new_col_letter].width = max(12, w_px / 7 + 2)
        return 1

    qr_count = 0
    max_w_px = 0
    for r_idx, url in links:
        try:
            qr_bytes, (w_px, h_px) = generate_qr_bytes(url, box_size=box_size, border=border)
        except Exception:
            continue
        img = XLImage(io.BytesIO(qr_bytes))
        excel_row = r_idx + 1  # 1-indexed
        ws.add_image(img, f"{new_col_letter}{excel_row}")
        target_h = h_px * 0.75 + 4  # px -> pt, + a little padding
        current_h = ws.row_dimensions[excel_row].height
        if not current_h or current_h < target_h:
            ws.row_dimensions[excel_row].height = target_h
        max_w_px = max(max_w_px, w_px)
        qr_count += 1

    if qr_count:
        ws.column_dimensions[new_col_letter].width = max(12, max_w_px / 7 + 2)

    return qr_count


# ---- "Keep template": preserve original .xls formatting (font, fill, border,
# column widths, row heights, merged cells) when rebuilding a legacy .xls into
# .xlsx, instead of a bare values-only rebuild. Only needed for the .xls
# fallback — real .xlsx already keeps 100% of its original formatting since
# that path edits the loaded workbook in place. ----
_QR_BORDER_STYLE_MAP = {
    0: None, 1: 'thin', 2: 'medium', 3: 'dashed', 4: 'dotted',
    5: 'thick', 6: 'double', 7: 'hair', 8: 'mediumDashed',
    9: 'dashDot', 10: 'mediumDashDot', 11: 'dashDotDot',
    12: 'mediumDashDotDot', 13: 'slantDashDot',
}
_QR_HALIGN_MAP = {1: 'left', 2: 'center', 3: 'right', 4: 'fill', 5: 'justify', 6: 'centerContinuous'}
_QR_VALIGN_MAP = {0: 'top', 1: 'center', 2: 'bottom', 3: 'justify'}


def _xlrd_color_hex(color_map, idx):
    """xlrd colour index -> 'RRGGBB' hex, or None for an 'automatic/default'
    color (index not in the palette — leave openpyxl's own default alone)."""
    if idx is None:
        return None
    rgb = color_map.get(idx)
    return ('%02X%02X%02X' % rgb) if rgb else None


def _xlrd_border_side(style_code, colour_idx, color_map):
    if not style_code:
        return Side(style=None)
    style = _QR_BORDER_STYLE_MAP.get(style_code, 'thin')
    color_hex = _xlrd_color_hex(color_map, colour_idx)
    return Side(style=style, color=color_hex) if color_hex else Side(style=style)


def _build_openpyxl_style_from_xf(xf, font_rec, color_map):
    """One xlrd XF (+ its font) record -> (Font, PatternFill, Border, Alignment)."""
    font = Font(
        name=font_rec.name or 'Calibri',
        size=(font_rec.height / 20.0) if font_rec.height else 11,
        bold=bool(font_rec.bold),
        italic=bool(font_rec.italic),
        color=_xlrd_color_hex(color_map, font_rec.colour_index),
    )
    fill = PatternFill()
    if getattr(xf.background, 'fill_pattern', 0):
        bg_hex = _xlrd_color_hex(color_map, xf.background.pattern_colour_index)
        if bg_hex:
            fill = PatternFill(fill_type='solid', fgColor=bg_hex)
    border = Border(
        top=_xlrd_border_side(xf.border.top_line_style, xf.border.top_colour_index, color_map),
        bottom=_xlrd_border_side(xf.border.bottom_line_style, xf.border.bottom_colour_index, color_map),
        left=_xlrd_border_side(xf.border.left_line_style, xf.border.left_colour_index, color_map),
        right=_xlrd_border_side(xf.border.right_line_style, xf.border.right_colour_index, color_map),
    )
    alignment = Alignment(
        horizontal=_QR_HALIGN_MAP.get(xf.alignment.hor_align),
        vertical=_QR_VALIGN_MAP.get(xf.alignment.vert_align),
        wrap_text=bool(getattr(xf.alignment, 'wrap', False)),
    )
    return font, fill, border, alignment


def _apply_xls_styles(ws, sh, xlrd_book, n_rows, n_cols):
    """Layers font/fill/border/alignment + column widths + row heights +
    merged cells from an xlrd(formatting_info=True) sheet onto an already
    value-populated openpyxl worksheet ws — bounded to [0,n_rows)x[0,n_cols),
    the same real-data bounds pandas already determined. That bound matters:
    some real CCS .xls exports report tens of thousands of blank padding rows
    at the raw-file level (e.g. a stray format applied far past the real data)
    — capping the loop at pandas' trimmed shape keeps this fast and correct
    regardless of what the raw file claims."""
    color_map = xlrd_book.colour_map
    xf_list = xlrd_book.xf_list
    font_list = xlrd_book.font_list
    style_cache = {}

    r_bound = min(n_rows, sh.nrows)
    c_bound = min(n_cols, sh.ncols)

    for r in range(r_bound):
        for c in range(c_bound):
            try:
                xf_index = sh.cell_xf_index(r, c)
            except IndexError:
                continue
            if xf_index not in style_cache:
                try:
                    xf = xf_list[xf_index]
                    font_rec = font_list[xf.font_index]
                    style_cache[xf_index] = _build_openpyxl_style_from_xf(xf, font_rec, color_map)
                except Exception:
                    style_cache[xf_index] = None
            style = style_cache[xf_index]
            if style is None:
                continue
            font, fill, border, alignment = style
            wc = ws.cell(row=r + 1, column=c + 1)
            wc.font = font
            wc.fill = fill
            wc.border = border
            wc.alignment = alignment

    for c in range(c_bound):
        ci = sh.colinfo_map.get(c)
        if ci and ci.width:
            ws.column_dimensions[get_column_letter(c + 1)].width = ci.width / 256.0

    for r in range(r_bound):
        ri = sh.rowinfo_map.get(r)
        if ri and ri.height:
            ws.row_dimensions[r + 1].height = ri.height / 20.0

    for (rlo, rhi, clo, chi) in sh.merged_cells:
        if rlo < r_bound and clo < c_bound:
            try:
                ws.merge_cells(start_row=rlo + 1, end_row=min(rhi, r_bound),
                                start_column=clo + 1, end_column=min(chi, c_bound))
            except Exception:
                pass


_HYPERLINK_FORMULA_RE = re.compile(
    r'^=\s*HYPERLINK\(\s*"((?:[^"]|"")*)"\s*(?:[,;]\s*"((?:[^"]|"")*)"\s*)?\)\s*$',
    re.IGNORECASE,
)


def _flatten_hyperlink_formulas(ws, cached_ws=None):
    """=HYPERLINK("url","ชื่อ") -> ค่าเซลล์ = "ชื่อ" (ข้อความธรรมดา) + cell.hyperlink = url
    ชื่อที่ใช้: ค่าที่ Excel คำนวณเก็บไว้ (ถ้ามี) ไม่งั้นอ่านจากอาร์กิวเมนต์ตัวที่ 2 ในสูตร
    ถ้าสูตรซับซ้อนกว่านี้ (อ้างเซลล์อื่น ฯลฯ) และไม่มีค่าที่คำนวณไว้ -> ปล่อยสูตรเดิมไว้ ไม่เดา
    Returns จำนวนเซลล์ที่แปลง"""
    n = 0
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.lstrip().upper().startswith('=HYPERLINK(')):
                continue
            m = _HYPERLINK_FORMULA_RE.match(v.strip())
            url = m.group(1).replace('""', '"') if m else None
            name = m.group(2).replace('""', '"') if (m and m.group(2) is not None) else None
            cached = cached_ws[cell.coordinate].value if cached_ws is not None else None
            display = cached if cached not in (None, '') else (name if name is not None else url)
            if display is None:
                continue
            cell.value = display
            if url:
                cell.hyperlink = url
            n += 1
    return n


def process_qr_batch_file(file_bytes, selected_sheets=None, box_size=5, border=2):
    """
    Scans each selected sheet for a column containing links and returns a NEW
    workbook (bytes) with a 'QR Code' column appended to every sheet that has
    one — one embedded, natively-scannable QR image per row with a link.

    Never alters or removes any existing cell — pure append. A sheet with no
    detected link is left completely untouched (not even a header added).

    Returns (output_bytes, report): report is a list of dicts per sheet —
    {'sheet', 'link_col_letter' (or None), 'qr_count', 'total_rows'}.
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheet_names = xl.sheet_names
    sheets_to_process = selected_sheets if selected_sheets else all_sheet_names

    # High-fidelity path: real .xlsx loads straight into openpyxl, so every
    # other sheet, style, formula and column is preserved byte-for-byte and we
    # only ever append the new column. Legacy .xls can't be loaded (or written)
    # by openpyxl, so it falls back to a values-only rebuild — same as the rest
    # of this app already does for .xls uploads.
    try:
        out_wb = load_workbook(io.BytesIO(file_bytes))
        native = True
    except Exception:
        out_wb = None
        native = False

    report = []

    if native:
        # เซลล์ที่เป็นสูตร =HYPERLINK(...) (เช่นคอลัมน์ "ที่ตั้ง" ในไฟล์ OptimoRoute) -> เขียนชื่อจริงลงไป
        # เป็นข้อความธรรมดา + ใส่ลิงก์จริงไว้ที่เซลล์ (ยังกดได้) ทำก่อนสแกนหาลิงก์ เพื่อไม่ให้ข้อความสูตร
        # ที่มี https:// ถูกเข้าใจผิดว่าเป็นคอลัมน์ลิงก์
        try:
            cached_wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
        except Exception:
            cached_wb = None
        for ws_any in out_wb.worksheets:
            cached_ws = cached_wb[ws_any.title] if cached_wb is not None and ws_any.title in cached_wb.sheetnames else None
            _flatten_hyperlink_formulas(ws_any, cached_ws)

        for sn in sheets_to_process:
            if sn not in out_wb.sheetnames:
                continue
            ws = out_wb[sn]
            raw_rows = list(ws.iter_rows(values_only=True))
            link_col = find_link_column(raw_rows)

            if link_col is None:
                report.append({'sheet': sn, 'link_col_letter': None, 'qr_count': 0, 'total_rows': len(raw_rows)})
                continue

            qr_count = _embed_qr_column(ws, raw_rows, link_col, ws.max_column, box_size, border)
            report.append({
                'sheet': sn,
                'link_col_letter': get_column_letter(link_col + 1),
                'qr_count': qr_count,
                'total_rows': len(raw_rows),
            })

        buf = io.BytesIO()
        out_wb.save(buf)
        return buf.getvalue(), report

    else:
        # Legacy .xls: values still come from the same pandas path as before
        # (unchanged — already trims fake padding ranges and handles date/time
        # cells correctly). "Keep template": layer the original font/fill/
        # border/column-width/merge formatting back on top via a separate
        # xlrd(formatting_info=True) pass. If that pass fails for any reason,
        # degrade silently to the old values-only look rather than breaking
        # the actual QR feature over a styling problem.
        try:
            xlrd_book = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
        except Exception:
            xlrd_book = None

        out_wb = Workbook()
        out_wb.remove(out_wb.active)

        for sn in sheets_to_process:
            if sn not in all_sheet_names:
                continue
            df_raw = pd.read_excel(xl, sheet_name=sn, header=None)
            raw_rows = df_raw.values.tolist()

            ws = out_wb.create_sheet(title=sn[:31])
            for r_idx, row in enumerate(raw_rows):
                for c_idx, v in enumerate(row):
                    if pd.isna(v):
                        continue
                    ws.cell(row=r_idx + 1, column=c_idx + 1, value=v)

            if xlrd_book is not None:
                try:
                    _apply_xls_styles(ws, xlrd_book.sheet_by_name(sn), xlrd_book, df_raw.shape[0], df_raw.shape[1])
                except Exception:
                    pass

            link_col = find_link_column(raw_rows)
            if link_col is None:
                report.append({'sheet': sn, 'link_col_letter': None, 'qr_count': 0, 'total_rows': len(raw_rows)})
                continue

            qr_count = _embed_qr_column(ws, raw_rows, link_col, df_raw.shape[1], box_size, border)
            report.append({
                'sheet': sn,
                'link_col_letter': get_column_letter(link_col + 1),
                'qr_count': qr_count,
                'total_rows': len(raw_rows),
            })

        buf = io.BytesIO()
        out_wb.save(buf)
        return buf.getvalue(), report


with tab_qr:
    st.markdown("##### 🔗 สร้าง QR Code จากลิงก์")
    st.caption("ใส่ลิงก์อะไรก็ได้ — ลิงก์แอป, Google Sheets, เอกสาร ฯลฯ — ได้ QR Code กลับมาดาวน์โหลดได้ทันที")

    qr_link = st.text_input(
        "ลิงก์", label_visibility="collapsed",
        placeholder="https://example.com",
        key="qr_link_input",
    )

    with st.expander("⚙️ ปรับแต่ง (ไม่บังคับ)"):
        qc1, qc2 = st.columns(2)
        with qc1:
            qr_size = st.slider("ขนาด QR Code", min_value=5, max_value=20, value=10, key="qr_size")
            qr_fill = st.color_picker("สีลาย", value="#000000", key="qr_fill")
        with qc2:
            qr_border = st.slider("ขอบขาว (border)", min_value=1, max_value=10, value=4, key="qr_border")
            qr_back = st.color_picker("สีพื้นหลัง", value="#FFFFFF", key="qr_back")

    if qr_link and qr_link.strip():
        try:
            qr_bytes, qr_dims = generate_qr_bytes(
                qr_link.strip(),
                box_size=qr_size, border=qr_border,
                fill_color=qr_fill, back_color=qr_back,
            )
            st.write("")
            qr_img_col, qr_info_col = st.columns([1, 1])
            with qr_img_col:
                st.image(qr_bytes, width=280)
            with qr_info_col:
                st.markdown(f"""
                <div class="rcm-card rcm-card-success">
                    ✅ <b>สร้าง QR Code สำเร็จ</b>
                    <div class="rcm-status-sub" style="margin-top:6px;">ขนาดภาพ: {qr_dims[0]}×{qr_dims[1]} px</div>
                </div>
                """, unsafe_allow_html=True)
                st.download_button(
                    "⬇️ ดาวน์โหลด QR Code (PNG)",
                    data=qr_bytes,
                    file_name="qrcode.png",
                    mime="image/png",
                    use_container_width=True,
                    key="qr_download_btn",
                )
        except Exception as e:
            st.error(f"สร้าง QR Code ไม่สำเร็จ: {e}")
    else:
        st.markdown("""
        <div class="rcm-card" style="text-align:center; color:#6B84A6;">
            📎 ใส่ลิงก์ด้านบนเพื่อเริ่มสร้าง QR Code
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="rcm-glow-divider"></div>', unsafe_allow_html=True)

    # ============================================================
    # BATCH: อ่านลิงก์จากไฟล์ที่อัพโหลด → ได้ไฟล์เดิมที่ฝัง QR Code กลับมา
    # ============================================================
    st.markdown("##### 📁 สร้าง QR Code จากไฟล์ (Batch)")
    st.caption("อัพโหลดไฟล์ที่มีลิงก์ Google Maps อยู่แล้ว (เช่นไฟล์ route ที่มีลิงก์เส้นทางท้ายชีท หรือ Master Data ที่มีลิงก์ต่อแถว) ระบบจะหาคอลัมน์ที่มีลิงก์ให้เอง แล้วฝัง QR Code กลับเข้าไปในคอลัมน์ใหม่ท้ายไฟล์ — ข้อมูลเดิมไม่ถูกแก้หรือลบแม้แต่แถวเดียว")

    qr_file = st.file_uploader("อัพโหลดไฟล์", type=['xls', 'xlsx'], label_visibility="collapsed", key="qr_batch_uploader")

    if qr_file is not None:
        qr_file_bytes = qr_file.read()

        try:
            qr_xl_peek = pd.ExcelFile(io.BytesIO(qr_file_bytes))
            qr_all_sheet_names = qr_xl_peek.sheet_names
        except Exception as e:
            st.error(f"เปิดไฟล์ไม่ได้: {e}")
            st.stop()

        st.markdown("###### 📑 เลือก Sheet ที่ต้องการสแกนหาลิงก์")
        qr_sheet_cols = st.columns(min(len(qr_all_sheet_names), 4))
        qr_selected_sheets = []
        for i, sn in enumerate(qr_all_sheet_names):
            with qr_sheet_cols[i % len(qr_sheet_cols)]:
                checked = st.checkbox(sn, value=True, key=f"qrbatch_sheet_{sn}")
                if checked:
                    qr_selected_sheets.append(sn)

        if not qr_selected_sheets:
            st.warning("⚠️ กรุณาเลือกอย่างน้อย 1 Sheet")
            st.stop()

        st.write("")
        qr_batch_clicked = st.button("🔍 สแกนหาลิงก์ + สร้าง QR Code", type="primary", use_container_width=True, key="qr_batch_btn")

        if qr_batch_clicked:
            with st.spinner("🔍 กำลังสแกนหาลิงก์และสร้าง QR Code..."):
                try:
                    qr_out_bytes, qr_report = process_qr_batch_file(qr_file_bytes, qr_selected_sheets)
                    st.session_state['last_qr_batch_bytes'] = qr_out_bytes
                    st.session_state['last_qr_batch_report'] = qr_report
                    st.session_state['last_qr_batch_filename'] = qr_file.name
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")
                    st.exception(e)

        if 'last_qr_batch_report' in st.session_state and st.session_state.get('last_qr_batch_filename') == qr_file.name:
            qr_report = st.session_state['last_qr_batch_report']
            qr_out_bytes = st.session_state['last_qr_batch_bytes']
            total_qr = sum(r['qr_count'] for r in qr_report)

            if total_qr > 0:
                st.markdown(f"""
                <div class="rcm-card rcm-card-success">
                    🎉 <b>เสร็จแล้ว!</b> สร้าง QR Code ได้ {total_qr:,} อัน
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="rcm-card rcm-card-warn">
                    ⚠️ ไม่พบลิงก์ในไฟล์นี้เลย — ลองเช็คว่ามีคอลัมน์ที่เก็บลิงก์ Google Maps อยู่จริงไหม
                </div>
                """, unsafe_allow_html=True)

            qr_report_cols = st.columns(len(qr_report))
            for col, r in zip(qr_report_cols, qr_report):
                with col:
                    if r['link_col_letter'] is None:
                        stat_box(f"SHEET: {r['sheet']}", "ไม่พบลิงก์", f"{r['total_rows']:,} แถว")
                    else:
                        stat_box(f"SHEET: {r['sheet']}", f"{r['qr_count']:,} QR", f"พบลิงก์ในคอลัมน์ {r['link_col_letter']}")

            if total_qr > 0:
                st.write("")
                qr_batch_base_name = qr_file.name.rsplit('.', 1)[0]
                st.download_button(
                    "⬇️ ดาวน์โหลดไฟล์ที่ฝัง QR Code แล้ว (Excel)",
                    data=qr_out_bytes,
                    file_name=f"{qr_batch_base_name}_with_qr.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="qr_batch_download_btn",
                )


# ============================================================
# TAB 4: FOR BKK — Sheet1 -> ใบงานต่อออเดอร์ (แยก Retail และ InterExp/B&W ออกเป็นชีทของตัวเอง)
# ============================================================
BKK_COLUMNS = ['Order Date', 'OrderNo', 'Customers Code', 'Cust Type', 'Customer Name', 'ShipTo Name',
               'ShipTo.Address', 'Route', 'Remark', 'เวลาเข้า', 'เวลาออก', 'Latitude', 'Longitude']
# ความกว้างคอลัมน์ลอกจากใบงาน Sheet2 ที่ทีมใช้อยู่จริง (A..K) + Lat/Lon
BKK_COL_WIDTHS = [9.9, 39.1, 14.4, 9.1, 53.3, 53.7, 40.1, 19.4, 27.4, 7.7, 8.6, 12.5, 12.5]


def build_channel_map(df):
    """Customer Master ทั้งตาราง (ไม่ตัดแถวที่ยังไม่มีพิกัด) -> {ShipTo Name ที่ normalize แล้ว: Channel}
    Channel คือคอลัมน์ C ของ Master; ถ้าชื่อเดียวกันมีหลายแถว ใช้แถวแรกที่เจอ"""
    ch_col = next((c for c in df.columns if str(c).strip().lower() == 'channel'), None)
    if ch_col is None:
        ch_col = df.columns[2]
    out = {}
    for name, ch in zip(df['Ship To Name'], df[ch_col]):
        key = normalize(name)
        if key and key not in out and not _cell_blank(ch):
            out[key] = str(ch).strip()
    return out


@st.cache_data(ttl=300)
def load_master_channels(url):
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    if 'Ship To Name' not in df.columns:
        raise ValueError("ไม่พบคอลัมน์ Ship To Name ใน Master")
    return build_channel_map(df)


def _bkk_clean_key(v):
    """975484.0 -> 975484, ' 5 ' -> '5', 1.0 -> 1 (ให้ OrderNo / Route เทียบกันได้ตรงๆ)"""
    if isinstance(v, float) and v == int(v):
        return int(v)
    if isinstance(v, str):
        s = v.strip()
        return int(s) if s.isdigit() else s
    return v


def read_bkk_orders(file_bytes):
    """อ่าน Sheet1 เท่านั้น -> 1 แถวต่อ 1 OrderNo (แถวแรกที่เจอของออเดอร์นั้น) ตามลำดับเดิมในไฟล์
    ตัดแถวว่างทิ้ง Returns (orders_df, info)"""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = 'Sheet1' if 'Sheet1' in xl.sheet_names else xl.sheet_names[0]
    raw = xl.parse(sheet, header=None).values.tolist()
    header_row, c_order = find_header_row(raw, ['orderno', 'order no', 'order no.'])
    if header_row is None:
        raise ValueError(f'ไม่พบคอลัมน์ OrderNo ในชีท "{sheet}"')
    header = [v.strip().lower() if isinstance(v, str) else '' for v in raw[header_row]]

    def col(*names, prefix=None):
        for i, h in enumerate(header):
            if h in names:
                return i
        if prefix:
            for i, h in enumerate(header):
                if h.startswith(prefix):
                    return i
        return None

    c_date = col('order date')
    c_code = col('cust code', 'cust id', 'cust. id', 'cust.id', 'customers code', 'customer code')
    c_type = col('cust type', 'customer type')
    c_cname = col('cust name', 'cust. name', 'customer name')
    c_sname = col('shipto name', 'ship to name', 'ship-to name')
    c_addr = col('shipto.address', prefix='shipto.address')
    if c_addr is None:
        c_addr = col('shipto.c')
    c_route = col('route')

    def get(row, c):
        if c is None or c >= len(row) or _cell_blank(row[c]):
            return None
        v = row[c]
        return v.strip() if isinstance(v, str) else v

    rows, seen, multi_route = [], {}, {}
    line_count = no_order_rows = 0
    for row in raw[header_row + 1:]:
        if all(_cell_blank(v) for v in row):
            continue  # แถวว่าง
        if _cell_blank(row[c_order]):
            no_order_rows += 1  # มีข้อมูลแต่ไม่มีเลขออเดอร์ -> นับไว้รายงาน
            continue
        line_count += 1
        order_no = _bkk_clean_key(row[c_order])
        route = _bkk_clean_key(get(row, c_route))
        if order_no in seen:
            first = seen[order_no]
            if route != first['Route'] and route is not None and first['Route'] is not None:
                multi_route.setdefault(order_no, [first['Route']])
                if route not in multi_route[order_no]:
                    multi_route[order_no].append(route)
            # แถวแรกของออเดอร์นี้มีช่องไหนว่าง แต่แถวถัดไปของออเดอร์เดียวกันมีค่า -> เติมให้ (ไม่ทับค่าที่มีอยู่)
            for fld, c in (('Order Date', c_date), ('Customers Code', c_code), ('Cust Type', c_type),
                           ('Customer Name', c_cname), ('ShipTo Name', c_sname), ('ShipTo.Address', c_addr),
                           ('Route', c_route)):
                if first[fld] is None:
                    v = get(row, c)
                    if v is not None:
                        first[fld] = _bkk_clean_key(v) if fld == 'Route' else v
            continue
        rec = {
            'Order Date': get(row, c_date), 'OrderNo': order_no, 'Customers Code': get(row, c_code),
            'Cust Type': get(row, c_type), 'Customer Name': get(row, c_cname),
            'ShipTo Name': get(row, c_sname), 'ShipTo.Address': get(row, c_addr), 'Route': route,
        }
        seen[order_no] = rec
        rows.append(rec)

    orders = pd.DataFrame(rows, columns=BKK_COLUMNS[:8])
    info = {'sheet': sheet, 'lines': line_count, 'orders': len(orders), 'no_order_rows': no_order_rows,
            'multi_route': multi_route, 'has_shipto_col': c_sname is not None}
    return orders, info


def _bkk_route_sort_key(route):
    """Route ตัวเลขมาก่อนเรียงตามค่า (1, 2, 3, 11, 12) แล้วตามด้วย Route ที่เป็นตัวหนังสือ"""
    if isinstance(route, (int, float)) and not pd.isna(route):
        return (0, float(route), '')
    return (1, 0.0, '' if route is None else str(route))


BKK_CARRIER_SHEET = 'InterExp-B&W'
BKK_NOTYPE_SHEET = 'No Cust Type'

# ลูกค้าที่ในระบบเป็น R (Retail) แต่ให้ถือเป็น F (Food Service) เสมอ — เพิ่มชื่อในลิสต์นี้ได้
BKK_FORCE_FOOD_SERVICE = ['FMR-for RT Sale Dept.']


def _bkk_name_key(v):
    """เทียบชื่อแบบไม่สนตัวพิมพ์ เว้นวรรค จุด ขีด: 'FMR- for RT Sale Dept' == 'FMR-for RT Sale Dept.'"""
    return re.sub(r'[^A-Z0-9ก-๙]', '', str(v).upper()) if v is not None else ''


def _bkk_is_carrier_route(route):
    """Route ที่ส่งผ่านขนส่งภายนอก: InterExp และ IE,B&W (เขียนได้หลายแบบ เช่น 'IE B&W', 'IE,BW', 'B&W')"""
    key = re.sub(r'[^A-Z0-9]', '', str(route).upper()) if route is not None else ''
    return key.startswith('INTEREXP') or key in ('IEBW', 'BW', 'IE')


def process_bkk_orders(orders, info, master_df, channel_map, dates=None):
    """แยก Retail / ไม่ใช่ Retail + เติม Lat/Lon (จับด้วย ShipTo Name แบบเดียวกับแท็บประมวลผลไฟล์)
    Retail ดูจาก Channel ใน Master (คอลัมน์ C) ตาม ShipTo Name; ถ้าไม่เจอชื่อใน Master เลย
    ใช้ Cust Type ในไฟล์แทน (R = Retail) แล้วนับไว้รายงาน
    ลำดับการแยกชีท:
      1) Route เป็น InterExp / IE,B&W -> ชีท InterExp-B&W เสมอ (ไม่ว่าจะเป็น Retail หรือไม่)
      2) ช่อง Cust Type ในไฟล์ว่าง -> ชีท No Cust Type (เก็บไว้ให้คนตรวจ ไม่เดาว่าเป็นประเภทไหน)
      3) ที่เหลือแบ่ง Retail / ไม่ใช่ Retail
    Returns (normal_df, retail_df, carrier_df, notype_df, info)"""
    df = orders.copy()
    if dates is not None:
        keep = set(str(d) for d in dates)
        df = df[df['Order Date'].astype(str).isin(keep)]

    by_ship = master_df.drop_duplicates('norm_ship').set_index('norm_ship')
    name_col = 'ShipTo Name' if info.get('has_shipto_col', True) else 'Customer Name'

    forced_keys = {_bkk_name_key(n) for n in BKK_FORCE_FOOD_SERVICE}
    lats, lons, is_retail, cust_types, by_cust_type, forced_fs = [], [], [], [], 0, 0
    for _, r in df.iterrows():
        key = normalize(r[name_col])
        ch = channel_map.get(key) if key else None
        if _bkk_name_key(r['ShipTo Name']) in forced_keys or _bkk_name_key(r['Customer Name']) in forced_keys:
            # ข้อยกเว้นที่กำหนดไว้: ถือเป็น Food Service เสมอ และเปลี่ยน Cust Type เป็น F
            forced_fs += 1
            is_retail.append(False)
            cust_types.append('F')
        elif ch is not None:
            is_retail.append(ch.strip().lower() == 'retail')
            cust_types.append(r['Cust Type'])
        else:
            by_cust_type += 1
            is_retail.append(str(r['Cust Type']).strip().upper() == 'R')
            cust_types.append(r['Cust Type'])
        if key and key in by_ship.index:
            hit = by_ship.loc[key]
            lats.append(hit['Latitude'])
            lons.append(hit['Longitude'])
        else:
            lats.append(None)
            lons.append(None)

    df['Cust Type'] = cust_types
    df['Remark'] = None
    df['เวลาเข้า'] = None
    df['เวลาออก'] = None
    df['Latitude'] = lats
    df['Longitude'] = lons
    df['_retail'] = is_retail
    df['_rk'] = df['Route'].map(_bkk_route_sort_key)
    df['_ck'] = df['Customer Name'].map(lambda v: '' if v is None else str(v))
    df['_ok'] = df['OrderNo'].map(lambda v: (0, v, '') if isinstance(v, int) else (1, 0, str(v)))
    df = df.sort_values(['_rk', '_ck', '_ok'], kind='stable')

    df['_carrier'] = df['Route'].map(_bkk_is_carrier_route)
    df['_notype'] = df['Cust Type'].map(_cell_blank) & ~df['_carrier']
    rest = ~df['_carrier'] & ~df['_notype']
    normal = df[rest & ~df['_retail']][BKK_COLUMNS].reset_index(drop=True)
    retail = df[rest & df['_retail']][BKK_COLUMNS].reset_index(drop=True)
    carrier = df[df['_carrier']][BKK_COLUMNS].reset_index(drop=True)
    notype = df[df['_notype']][BKK_COLUMNS].reset_index(drop=True)
    # คอลัมน์ Route ในไฟล์ผลลัพธ์ = Route จากคอลัมน์ A ของ Sheet1 (ออเดอร์ที่มีหลาย Route ใช้ Route แรกที่เจอ)
    out_info = dict(info)
    out_info.update({
        'selected_orders': len(df), 'normal': len(normal), 'retail': len(retail), 'carrier': len(carrier),
        'notype': len(notype), 'coords_notype': int(notype['Latitude'].notna().sum()),
        'carrier_notype': int((df['_carrier'] & df['Cust Type'].map(_cell_blank)).sum()),
        'carrier_retail': int((df['_carrier'] & df['_retail']).sum()),
        'retail_by_cust_type': by_cust_type, 'forced_fs': forced_fs,
        'coords_normal': int(normal['Latitude'].notna().sum()), 'coords_retail': int(retail['Latitude'].notna().sum()),
        'coords_carrier': int(carrier['Latitude'].notna().sum()),
    })
    return normal, retail, carrier, notype, out_info


def bkk_to_excel_bytes(normal_df, retail_df, carrier_df, notype_df=None):
    """ชีท Order / Retail / InterExp-B&W (+ No Cust Type ถ้ามีออเดอร์ที่ช่อง Cust Type ว่าง)
    หน้าตาเดียวกับใบงาน Sheet2: หัวกระดาษ ทะเบียนรถ/วันที่/เลขไมล์,
    หัวตารางแถว 4, ฟอนต์ Cordia New, เส้นขอบทุกช่อง + Lat/Lon ต่อท้าย"""
    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style='thin')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    f_top = Font(name='Cordia New', size=12, bold=True)
    f_head = Font(name='Cordia New', size=14, bold=True, underline='single')
    f_body = Font(name='Cordia New', size=14)
    center_cols = {'Cust Type', 'Route'}

    sheets = [('Order', normal_df), ('Retail', retail_df), (BKK_CARRIER_SHEET, carrier_df)]
    if notype_df is not None and len(notype_df):
        sheets.append((BKK_NOTYPE_SHEET, notype_df))
    for title, df in sheets:
        ws = wb.create_sheet(title=title)
        ws['B1'] = 'ทะเบียนรถ____________________'
        ws['F1'] = 'วันที่______________________'
        ws['B2'] = 'เลขไมล์ออกจากคลัง___________________________'
        ws['F2'] = 'เลขไมล์กลับถึงคลัง___________________________'
        for ref in ('B1', 'F1', 'B2', 'F2'):
            ws[ref].font = f_top
        for r in (1, 2, 3):
            ws.row_dimensions[r].height = 18.75

        for c, name in enumerate(BKK_COLUMNS, start=1):
            cell = ws.cell(row=4, column=c, value=name)
            cell.font = f_head
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[4].height = 21

        for i, rec in enumerate(df.itertuples(index=False), start=5):
            for c, (name, v) in enumerate(zip(BKK_COLUMNS, rec), start=1):
                if v is not None and isinstance(v, float) and pd.isna(v):
                    v = None
                cell = ws.cell(row=i, column=c, value=v)
                cell.font = f_body
                cell.border = border
                cell.alignment = Alignment(horizontal='center' if name in center_cols else 'left', vertical='center')
            ws.row_dimensions[i].height = 21.75

        for c, w in enumerate(BKK_COL_WIDTHS, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.freeze_panes = 'A5'
        ws.print_title_rows = '1:4'
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


with tab_bkk:
    st.markdown("##### 🏙️ For BKK — ทำใบงานจาก Sheet1")
    st.caption("อัพโหลดไฟล์ route (ใช้เฉพาะ Sheet1) ระบบจะตัดแถวว่าง ยุบให้เหลือ 1 แถวต่อ 1 OrderNo "
               "แยก Route InterExp / IE,B&W ไปชีทหนึ่ง แยก Retail ไปอีกชีท (ดู Channel ใน Master จาก ShipTo Name) "
               "แล้วเติม Lat/Lon ให้ ได้ไฟล์เดียว: Order, Retail, InterExp-B&W หน้าตาเหมือนใบงาน Sheet2 "
               "(ออเดอร์ที่ช่อง Cust Type ว่างจะถูกเก็บไว้ในชีท No Cust Type ไม่ลบทิ้ง)")

    bkk_file = st.file_uploader("อัพโหลดไฟล์ route", type=['xls', 'xlsx'], label_visibility="collapsed", key="bkk_uploader")

    if bkk_file is not None:
        bkk_bytes = bkk_file.read()
        try:
            bkk_orders, bkk_info = read_bkk_orders(bkk_bytes)
        except Exception as e:
            st.error(f"อ่านไฟล์ไม่ได้: {e}")
            st.stop()

        bkk_dates = sorted({str(d) for d in bkk_orders['Order Date'] if d is not None})
        st.markdown(f"""
        <div class="rcm-card">
            📄 ชีท <b>{bkk_info['sheet']}</b>: {bkk_info['lines']:,} รายการสินค้า → <b>{bkk_info['orders']:,} ออเดอร์</b> (OrderNo ไม่ซ้ำ)
        </div>
        """, unsafe_allow_html=True)

        if len(bkk_dates) > 1:
            bkk_sel_dates = st.multiselect(
                "Order Date ที่จะเอา (ค่าเริ่มต้น = ทุกวันที่มีใน Sheet1)",
                options=bkk_dates, default=bkk_dates, key="bkk_dates",
            )
        else:
            bkk_sel_dates = bkk_dates

        if not bkk_sel_dates:
            st.warning("⚠️ กรุณาเลือก Order Date อย่างน้อย 1 วัน")
            st.stop()

        bkk_clicked = st.button("🚀 ทำใบงาน For BKK", type="primary", use_container_width=True, key="bkk_btn")

        if bkk_clicked:
            with st.spinner("🛰️ กำลังแยก Retail และเติมพิกัด..."):
                try:
                    bkk_channels = load_master_channels(MASTER_SHEET_CSV_URL)
                    bkk_normal, bkk_retail, bkk_carrier, bkk_notype, bkk_result = process_bkk_orders(
                        bkk_orders, bkk_info, master_df, bkk_channels,
                        dates=bkk_sel_dates if len(bkk_dates) > 1 else None,
                    )
                    st.session_state['bkk_xlsx'] = bkk_to_excel_bytes(bkk_normal, bkk_retail, bkk_carrier, bkk_notype)
                    st.session_state['bkk_result'] = bkk_result
                    st.session_state['bkk_filename'] = bkk_file.name
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")
                    st.exception(e)

        if 'bkk_result' in st.session_state and st.session_state.get('bkk_filename') == bkk_file.name:
            res = st.session_state['bkk_result']
            st.markdown(f"""
            <div class="rcm-card rcm-card-success">
                🎯 <b>เสร็จแล้ว!</b> {res['selected_orders']:,} ออเดอร์ → Order {res['normal']:,} + Retail {res['retail']:,} + InterExp/B&W {res['carrier']:,} + ไม่มี Cust Type {res['notype']:,}
            </div>
            """, unsafe_allow_html=True)

            bc1, bc2, bc3, bc4 = st.columns(4)
            with bc1:
                stat_box("SHEET: Order", f"{res['normal']:,}",
                          f"มีพิกัด {res['coords_normal']:,} / {res['normal']:,}")
            with bc2:
                stat_box("SHEET: Retail", f"{res['retail']:,}",
                          f"มีพิกัด {res['coords_retail']:,} / {res['retail']:,}")
            with bc3:
                stat_box("SHEET: InterExp-B&W", f"{res['carrier']:,}",
                          f"มีพิกัด {res['coords_carrier']:,} / {res['carrier']:,} · เป็น Retail {res['carrier_retail']:,}")
            with bc4:
                stat_box("SHEET: No Cust Type", f"{res['notype']:,}",
                          f"มีพิกัด {res['coords_notype']:,} / {res['notype']:,}" if res['notype'] else "ไม่มี — ไม่สร้างชีทนี้")

            bkk_notes = []
            if res['retail_by_cust_type']:
                bkk_notes.append(f"{res['retail_by_cust_type']:,} ออเดอร์ไม่เจอ ShipTo Name ใน Master — แยก Retail จาก Cust Type ในไฟล์แทน (R = Retail)")
            if res['notype']:
                bkk_notes.append(f"{res['notype']:,} ออเดอร์ช่อง Cust Type ว่าง — แยกไว้ในชีท {BKK_NOTYPE_SHEET} ให้ตรวจเอง (ไม่ได้ลบ และไม่ได้เดาประเภท)")
            if res.get('carrier_notype'):
                bkk_notes.append(f"{res['carrier_notype']:,} ออเดอร์ในชีท {BKK_CARRIER_SHEET} ช่อง Cust Type ว่าง (อยู่ชีทนั้นตาม Route)")
            if res.get('forced_fs'):
                bkk_notes.append(f"{res['forced_fs']:,} ออเดอร์ของ {', '.join(BKK_FORCE_FOOD_SERVICE)} ถูกเปลี่ยนเป็น F (Food Service) ตามที่ตั้งไว้")
            if res['multi_route']:
                mr = ', '.join(f"{k} ({' / '.join(str(x) for x in v)})" for k, v in list(res['multi_route'].items())[:10])
                bkk_notes.append(f"{len(res['multi_route']):,} ออเดอร์มีมากกว่า 1 Route ใน Sheet1 — ใช้ Route แรกที่เจอในการแยกชีท: {mr}")
            if res['no_order_rows']:
                bkk_notes.append(f"{res['no_order_rows']:,} แถวมีข้อมูลแต่ไม่มี OrderNo — ไม่ได้เอามา")
            if bkk_notes:
                st.markdown('<div class="rcm-card rcm-card-warn">⚠️ <b>ข้อสังเกต</b><br>' + '<br>'.join('• ' + n for n in bkk_notes) + '</div>',
                            unsafe_allow_html=True)

            st.write("")
            st.download_button(
                "⬇️ ดาวน์โหลดใบงาน For BKK (Excel: Order + Retail + InterExp-B&W + No Cust Type)",
                data=st.session_state['bkk_xlsx'],
                file_name=f"{bkk_file.name.rsplit('.', 1)[0]}_For_BKK.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="bkk_download_btn",
            )
