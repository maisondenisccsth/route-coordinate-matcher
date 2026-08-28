"""
Route Coordinate Matcher — Streamlit Web App
==============================================
โยนไฟล์ route เข้าไป ได้พิกัดกลับมาทันที ผ่านเว็บแอปที่มีลิงก์ถาวร

วิธี deploy: ดูขั้นตอนใน DEPLOY_INSTRUCTIONS.md
"""

import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="Route Coordinate Matcher", page_icon="🚚", layout="wide")

# ============================================================
# THEME — Logistics / Transportation style
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

    html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }

    .stApp {
        background: linear-gradient(180deg, #0B1220 0%, #0F1B2E 100%);
    }

    /* Header band */
    .rcm-header {
        background: linear-gradient(90deg, #0F2A4A 0%, #143A63 60%, #0F2A4A 100%);
        border: 1px solid #1E3A5F;
        border-radius: 14px;
        padding: 22px 28px;
        margin-bottom: 22px;
        position: relative;
        overflow: hidden;
    }
    .rcm-header::after {
        content: "";
        position: absolute;
        right: -40px; top: -40px;
        width: 180px; height: 180px;
        background: radial-gradient(circle, rgba(255,140,66,0.20) 0%, transparent 70%);
    }
    .rcm-header::before {
        content: "";
        position: absolute;
        left: 0; bottom: 0;
        width: 100%; height: 3px;
        background: linear-gradient(90deg, #FF6B35, transparent 60%);
    }
    .rcm-eyebrow {
        font-family: 'JetBrains Mono', monospace;
        font-size: 11px;
        letter-spacing: 0.15em;
        color: #FF8C42;
        font-weight: 600;
        margin-bottom: 6px;
        text-transform: uppercase;
    }
    .rcm-title {
        font-size: 28px;
        font-weight: 800;
        color: #F4F7FB;
        margin: 0 0 4px 0;
        letter-spacing: -0.02em;
    }
    .rcm-subtitle {
        font-size: 14px;
        color: #8FA8C7;
        font-family: 'JetBrains Mono', monospace;
    }

    /* Cards */
    .rcm-card {
        background: #101B2E;
        border: 1px solid #1E3A5F;
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 16px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.15);
    }
    .rcm-card-success {
        border-color: #1F5C42;
        background: linear-gradient(90deg, rgba(31,92,66,0.16), rgba(16,27,46,0.4));
    }
    .rcm-card-warn {
        border-color: #7A5423;
        background: linear-gradient(90deg, rgba(122,84,35,0.16), rgba(16,27,46,0.4));
    }

    /* Stat boxes */
    .rcm-stat {
        background: #0D1729;
        border: 1px solid #1E3A5F;
        border-radius: 10px;
        padding: 14px 16px;
        text-align: center;
        transition: border-color 0.15s ease, transform 0.15s ease;
    }
    .rcm-stat:hover {
        border-color: #FF8C42;
        transform: translateY(-1px);
    }
    .rcm-stat-label {
        font-family: 'JetBrains Mono', monospace;
        font-size: 11px;
        color: #6B84A6;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 6px;
    }
    .rcm-stat-value {
        font-size: 22px;
        font-weight: 800;
        color: #F4F7FB;
    }
    .rcm-stat-sub {
        font-family: 'JetBrains Mono', monospace;
        font-size: 11px;
        color: #FF8C42;
        margin-top: 2px;
    }

    /* Buttons */
    .stDownloadButton button, .stButton button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        border: none !important;
    }
    .stDownloadButton button {
        background: linear-gradient(90deg, #FF8C42, #FF6B35) !important;
        color: #0B1220 !important;
    }

    /* File uploader */
    [data-testid="stFileUploaderDropzone"] {
        background: #0D1729 !important;
        border: 2px dashed #2A4A72 !important;
        border-radius: 12px !important;
    }

    /* Dataframe */
    [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

    hr { border-color: #1E3A5F !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# CONNECT TO GOOGLE SHEETS (ฐานข้อมูลถาวร)
# ============================================================
MASTER_SHEET_CSV_URL = st.secrets.get("MASTER_SHEET_CSV_URL", "")


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
# UI
# ============================================================
st.markdown("""
<div class="rcm-header">
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
    </div>
    """, unsafe_allow_html=True)
except Exception as e:
    st.error(f"โหลด Master Data ไม่สำเร็จ: {e}")
    st.stop()

with st.expander("🔍 ดูตัวอย่าง Master Data"):
    st.dataframe(master_df[['Cust ID', 'Ship To Name', 'Latitude', 'Longitude']].head(20), use_container_width=True)

st.divider()

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

                st.write("")
                dl_col1, dl_col2 = st.columns(2)

                with dl_col1:
                    output_bytes = to_excel_bytes(sheets)
                    base_name = uploaded_file.name.rsplit('.', 1)[0]
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
                        ⚠️ <b>มี {len(unmatched_df)} รายการที่ยังไม่มีพิกัด</b> — ใช้ไฟล์ที่ดาวน์โหลดด้านบนไปเพิ่มลงใน Master Data ได้เลย
                    </div>
                    """, unsafe_allow_html=True)
                    st.dataframe(unmatched_df, use_container_width=True, hide_index=True)
                    st.caption("💡 นำ Cust Code เหล่านี้ไปเพิ่มพิกัดใน Google Sheets Master Data แล้วรีเฟรชหน้านี้ (หรือรอ cache หมดอายุใน 5 นาที)")

            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {e}")
                st.exception(e)

st.divider()
st.caption(f"🛰️ Master Data sync (cache): {datetime.now().strftime('%Y-%m-%d %H:%M')} — รีเฟรชทุก 5 นาที")
