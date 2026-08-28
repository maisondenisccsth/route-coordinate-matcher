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
from datetime import datetime

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
        font-size: 0 !important;
        padding: 9px 20px !important;
    }
    [data-testid="stFileUploaderDropzone"] button::after {
        content: "↑  Upload File";
        font-size: 13.5px;
        font-weight: 600;
        color: #F7FAFD;
        font-family: 'Inter', sans-serif;
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
</style>
""", unsafe_allow_html=True)

MASTER_SHEET_CSV_URL = st.secrets.get("MASTER_SHEET_CSV_URL", "")


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
    required = ['Cust ID', 'Ship To Name', 'Latitude', 'Longitude']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"ไม่พบคอลัมน์: {', '.join(missing)} กรุณาตรวจสอบหัวตารางใน Google Sheets")
    df = df.dropna(subset=['Cust ID', 'Latitude', 'Longitude'])
    df['norm_ship'] = df['Ship To Name'].apply(normalize)
    df['norm_code'] = df['Cust ID'].astype(str).str.strip().str.upper()
    return df


def find_header_row(rows, keys):
    for r in range(min(len(rows), 10)):
        row = rows[r]
        for c, v in enumerate(row):
            if isinstance(v, str) and v.strip().lower() in keys:
                return r, c
    return None, None


def process_route_file(file_bytes, master_df, selected_sheets=None):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    output_sheets = {}
    report = []
    unmatched = []

    by_ship = master_df.drop_duplicates('norm_ship').set_index('norm_ship')
    by_code = master_df.drop_duplicates('norm_code').set_index('norm_code')

    sheet_list = selected_sheets if selected_sheets is not None else xl.sheet_names

    for sheet_name in sheet_list:
        raw = xl.parse(sheet_name, header=None).values.tolist()
        header_row, cust_col = find_header_row(raw, ['cust code', 'cust id'])

        if header_row is None:
            output_sheets[sheet_name] = xl.parse(sheet_name, header=None)
            report.append({'sheet': sheet_name, 'skipped': True, 'matched': 0, 'total': 0})
            continue

        header = raw[header_row]
        ship_col = None
        for i, v in enumerate(header):
            if isinstance(v, str) and 'ship' in v.lower() and 'name' in v.lower():
                ship_col = i
                break

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
            ship_val = row[ship_col] if (ship_col is not None and ship_col < len(row)) else None

            hit = None
            if ship_val:
                key = normalize(ship_val)
                if key in by_ship.index:
                    hit = by_ship.loc[key]
                    via_ship += 1
            if hit is None:
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

    return output_sheets, report, unmatched


def to_excel_bytes(sheets_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=name[:31], header=False, index=False)
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

if not MASTER_SHEET_CSV_URL:
    st.error("⚠️ ยังไม่ได้ตั้งค่า MASTER_SHEET_CSV_URL — ดูวิธีตั้งค่าใน DEPLOY_INSTRUCTIONS.md")
    st.stop()

try:
    master_df = load_master_data(MASTER_SHEET_CSV_URL)
    st.markdown(f"""
    <div class="rcm-status-card">
        <div class="rcm-status-left">
            <div class="rcm-status-icon">✅</div>
            <div>
                <div class="rcm-status-title">เชื่อมต่อฐานข้อมูลสำเร็จ</div>
                <div class="rcm-status-sub">{len(master_df):,} รายชื่อลูกค้าในระบบ</div>
            </div>
        </div>
        <div class="rcm-status-right">
            <div class="rcm-status-right-icon">🗃️</div>
            <div class="rcm-status-pill"><span class="rcm-status-dot"></span>Connected</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
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
    st.dataframe(master_df[['Cust ID', 'Ship To Name', 'Latitude', 'Longitude']].head(20), use_container_width=True)

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

    st.write("")
    process_clicked = st.button("🚀 ประมวลผล Sheet ที่เลือก", type="primary", use_container_width=True)

    if process_clicked:
        with st.spinner("🛰️ กำลังจับคู่พิกัด..."):
            try:
                sheets, report, unmatched = process_route_file(file_bytes, master_df, selected_sheets)
                st.session_state['last_sheets'] = sheets
                st.session_state['last_report'] = report
                st.session_state['last_unmatched'] = unmatched
                st.session_state['last_filename'] = uploaded_file.name
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {e}")
                st.exception(e)

    if 'last_sheets' in st.session_state and st.session_state.get('last_filename') == uploaded_file.name:
        sheets = st.session_state['last_sheets']
        report = st.session_state['last_report']
        unmatched = st.session_state['last_unmatched']

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
        dl_col1, dl_col2 = st.columns(2)

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

        if not unmatched_df.empty:
            st.write("")
            st.markdown(f"""
            <div class="rcm-card rcm-card-warn">
                ⚠️ <b>มี {len(unmatched_df)} รายการที่ยังไม่มีพิกัด</b> — ใช้ไฟล์ที่ดาวน์โหลดด้านบนไปเพิ่มลงใน Master Data ได้เลย (ที่ Google Sheets โดยตรง)
            </div>
            """, unsafe_allow_html=True)
            st.dataframe(unmatched_df, use_container_width=True, hide_index=True)

st.markdown('<div class="rcm-glow-divider"></div>', unsafe_allow_html=True)
st.caption(f"🛰️ Master Data sync (cache): {datetime.now().strftime('%Y-%m-%d %H:%M')} — รีเฟรชทุก 5 นาที")
