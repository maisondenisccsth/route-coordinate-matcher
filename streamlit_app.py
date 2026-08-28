"""
Route Coordinate Matcher — Streamlit Web App (v3)
====================================================
โยนไฟล์ route เข้าไป ได้พิกัดกลับมาทันที + แก้ไข Master Data ในแอปได้เลย

วิธี deploy: ดู DEPLOY_INSTRUCTIONS.md
วิธีตั้งค่าแก้ไข Master Data: ดู SERVICE_ACCOUNT_SETUP.md
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

    /* ---------- Header ---------- */
    .rcm-header {
        background: linear-gradient(120deg, #0D1F38 0%, #123056 45%, #0D1F38 100%);
        border: 1px solid #204A78;
        border-radius: 18px;
        padding: 30px 34px;
        margin-bottom: 26px;
        position: relative;
        overflow: hidden;
        box-shadow: 0 10px 40px rgba(0,0,0,0.35), inset 0 1px 0 rgba(255,255,255,0.04);
    }
    .rcm-header::after {
        content: "";
        position: absolute;
        right: -60px; top: -60px;
        width: 260px; height: 260px;
        background: radial-gradient(circle, rgba(255,140,66,0.22) 0%, transparent 70%);
    }
    .rcm-header::before {
        content: "";
        position: absolute;
        left: 0; bottom: 0; width: 100%; height: 3px;
        background: linear-gradient(90deg, #FF6B35 0%, #FFB347 35%, transparent 75%);
    }
    .rcm-road {
        position: absolute; right: 34px; top: 50%; transform: translateY(-50%);
        font-size: 64px; opacity: 0.10; line-height: 1;
    }
    .rcm-eyebrow {
        font-family: 'JetBrains Mono', monospace;
        font-size: 11.5px;
        letter-spacing: 0.18em;
        color: #FF9B5C;
        font-weight: 700;
        margin-bottom: 8px;
        text-transform: uppercase;
    }
    .rcm-title {
        font-size: 32px;
        font-weight: 800;
        color: #F7FAFD;
        margin: 0 0 6px 0;
        letter-spacing: -0.02em;
    }
    .rcm-subtitle {
        font-size: 14.5px;
        color: #9FB8D9;
        font-family: 'JetBrains Mono', monospace;
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

    /* ---------- Tabs ---------- */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background: #0C1728;
        padding: 6px;
        border-radius: 12px;
        border: 1px solid #1E3A5F;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        color: #8FA8C7;
        font-weight: 600;
        padding: 8px 18px;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #1E3A5F, #234870) !important;
        color: #F7FAFD !important;
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

    /* ---------- File uploader ---------- */
    [data-testid="stFileUploaderDropzone"] {
        background: linear-gradient(180deg, #0F1D32, #0A1526) !important;
        border: 2px dashed #2A4A72 !important;
        border-radius: 14px !important;
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
MASTER_SHEET_ID = st.secrets.get("MASTER_SHEET_ID", "")
MASTER_SHEET_TAB = st.secrets.get("MASTER_SHEET_TAB", "Sheet1")
HAS_SERVICE_ACCOUNT = "gcp_service_account" in st.secrets


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


def get_gsheet_client():
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)


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
    <div class="rcm-road">🛣️</div>
    <div class="rcm-eyebrow">🚚 FLEET OPERATIONS · GEO-MATCHING SYSTEM</div>
    <div class="rcm-title">Route Coordinate Matcher</div>
    <div class="rcm-subtitle">โยนไฟล์ route → ได้พิกัดกลับมาทันที</div>
</div>
""", unsafe_allow_html=True)

if not MASTER_SHEET_CSV_URL:
    st.error("⚠️ ยังไม่ได้ตั้งค่า MASTER_SHEET_CSV_URL — ดูวิธีตั้งค่าใน DEPLOY_INSTRUCTIONS.md")
    st.stop()

try:
    master_df = load_master_data(MASTER_SHEET_CSV_URL)
    st.markdown(f"""
    <div class="rcm-card rcm-card-success">
        ✅ <b>เชื่อมต่อฐานข้อมูลสำเร็จ</b> — {len(master_df):,} รายชื่อลูกค้าในระบบ
        {' · 🔓 แก้ไขได้โดยตรง' if HAS_SERVICE_ACCOUNT else ' · 🔒 อ่านอย่างเดียว (ยังไม่ได้ตั้งค่า Service Account)'}
    </div>
    """, unsafe_allow_html=True)
except Exception as e:
    st.error(f"โหลด Master Data ไม่สำเร็จ: {e}")
    st.stop()

tab_process, tab_master = st.tabs(["📤  ประมวลผลไฟล์", "🗄️  ฐานข้อมูลลูกค้า"])

# ============================================================
# TAB 1: PROCESS FILES
# ============================================================
with tab_process:
    uploaded_file = st.file_uploader("📦 ลากไฟล์ route มาวางที่นี่ หรือคลิกเพื่อเลือกไฟล์", type=['xls', 'xlsx'])

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

            base_name = uploaded_file.name.rsplit('.', 1)[0]
            dl_col1, dl_col2 = st.columns(2)

            with dl_col1:
                output_bytes = to_excel_bytes(sheets)
                st.download_button(
                    "⬇️ ดาวน์โหลดไฟล์พร้อมพิกัด",
                    data=output_bytes,
                    file_name=f"{base_name}_with_coordinates.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
                    ⚠️ <b>มี {len(unmatched_df)} รายการที่ยังไม่มีพิกัด</b> — ใช้ไฟล์ที่ดาวน์โหลดด้านบนไปเพิ่มลงใน Master Data ได้เลย (แท็บ "ฐานข้อมูลลูกค้า")
                </div>
                """, unsafe_allow_html=True)
                st.dataframe(unmatched_df, use_container_width=True, hide_index=True)

# ============================================================
# TAB 2: MASTER DATA MANAGEMENT
# ============================================================
with tab_master:
    if not HAS_SERVICE_ACCOUNT or not MASTER_SHEET_ID:
        st.markdown("""
        <div class="rcm-card rcm-card-warn">
            🔒 <b>ยังแก้ไขข้อมูลในนี้ไม่ได้</b> — ต้องตั้งค่า Google Service Account ก่อน
            ดูวิธีตั้งค่าใน <code>SERVICE_ACCOUNT_SETUP.md</code><br><br>
            ระหว่างนี้แก้ไข Master Data ได้ที่ Google Sheets โดยตรง แล้วรีเฟรชหน้านี้
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(master_df[['Cust ID', 'Ship To Name', 'Latitude', 'Longitude']], use_container_width=True, hide_index=True)
    else:
        st.markdown("##### 🔍 ค้นหาลูกค้า")
        search_term = st.text_input("ค้นหาด้วย Cust ID หรือ Ship To Name", label_visibility="collapsed",
                                     placeholder="พิมพ์เพื่อค้นหา...")

        display_df = master_df[['Cust ID', 'Ship To Name', 'Latitude', 'Longitude']]
        if search_term:
            mask = (
                display_df['Cust ID'].astype(str).str.contains(search_term, case=False, na=False) |
                display_df['Ship To Name'].astype(str).str.contains(search_term, case=False, na=False)
            )
            display_df = display_df[mask]

        st.dataframe(display_df, use_container_width=True, hide_index=True, height=280)
        st.caption(f"แสดง {len(display_df):,} จาก {len(master_df):,} รายการทั้งหมด")

        st.markdown('<div class="rcm-glow-divider"></div>', unsafe_allow_html=True)
        st.markdown("##### ➕ เพิ่ม / แก้ไขลูกค้า")
        st.caption("ถ้า Cust ID + Ship To Name ตรงกับที่มีอยู่แล้ว ระบบจะอัพเดตพิกัดแทนที่ของเดิม")

        with st.form("add_customer_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                new_code = st.text_input("Cust ID")
                new_lat = st.text_input("Latitude")
            with c2:
                new_ship = st.text_input("Ship To Name")
                new_lon = st.text_input("Longitude")

            submitted = st.form_submit_button("💾 บันทึกเข้า Master Data", type="primary", use_container_width=True)

            if submitted:
                if not new_code or not new_ship:
                    st.error("กรุณากรอก Cust ID และ Ship To Name")
                else:
                    try:
                        lat_val = float(new_lat)
                        lon_val = float(new_lon)
                    except ValueError:
                        st.error("Latitude/Longitude ต้องเป็นตัวเลข")
                        st.stop()

                    try:
                        gc = get_gsheet_client()
                        sh = gc.open_by_key(MASTER_SHEET_ID)
                        ws = sh.worksheet(MASTER_SHEET_TAB)
                        all_values = ws.get_all_values()
                        header = all_values[0]
                        id_col = header.index('Cust ID')
                        ship_col = header.index('Ship To Name')
                        lat_col = header.index('Latitude')
                        lon_col = header.index('Longitude')

                        target_row = None
                        for i, row in enumerate(all_values[1:], start=2):
                            if (len(row) > max(id_col, ship_col) and
                                row[id_col].strip().upper() == new_code.strip().upper() and
                                normalize(row[ship_col]) == normalize(new_ship)):
                                target_row = i
                                break

                        if target_row:
                            ws.update_cell(target_row, lat_col + 1, lat_val)
                            ws.update_cell(target_row, lon_col + 1, lon_val)
                            st.success(f"✅ อัพเดตพิกัดของ {new_code} / {new_ship} เรียบร้อย")
                        else:
                            new_row = [''] * len(header)
                            new_row[id_col] = new_code
                            new_row[ship_col] = new_ship
                            new_row[lat_col] = lat_val
                            new_row[lon_col] = lon_val
                            ws.append_row(new_row)
                            st.success(f"✅ เพิ่ม {new_code} / {new_ship} เข้า Master Data เรียบร้อย")

                        load_master_data.clear()
                        st.rerun()

                    except Exception as e:
                        st.error(f"บันทึกไม่สำเร็จ: {e}")

st.markdown('<div class="rcm-glow-divider"></div>', unsafe_allow_html=True)
st.caption(f"🛰️ Master Data sync (cache): {datetime.now().strftime('%Y-%m-%d %H:%M')} — รีเฟรชทุก 5 นาที")
