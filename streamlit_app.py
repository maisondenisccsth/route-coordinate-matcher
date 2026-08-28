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

st.set_page_config(page_title="Route Coordinate Matcher", page_icon="📍", layout="wide")

# ============================================================
# CONNECT TO GOOGLE SHEETS (ฐานข้อมูลถาวร)
# ============================================================
# ใช้ Google Sheets เป็นที่เก็บ Master Data ถาวร ผ่าน public CSV export link
# วิธีตั้งค่า: ดูใน DEPLOY_INSTRUCTIONS.md ขั้นตอน "เชื่อม Google Sheets"

MASTER_SHEET_CSV_URL = st.secrets.get("MASTER_SHEET_CSV_URL", "")


def normalize(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ''
    return re.sub(r'\s+', ' ', str(s).strip().upper().replace('.', '').replace(',', ''))


@st.cache_data(ttl=300)  # cache 5 นาที กันโหลดบ่อยเกินไป
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


def process_route_file(file_bytes, master_df):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    output_sheets = {}
    report = []
    unmatched = []

    by_ship = master_df.drop_duplicates('norm_ship').set_index('norm_ship')
    by_code = master_df.drop_duplicates('norm_code').set_index('norm_code')

    for sheet_name in xl.sheet_names:
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
                unmatched.append({'Cust Code': code, 'Ship To Name': ship_val or '', 'Sheet': sheet_name})
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


# ============================================================
# UI
# ============================================================
st.title("📍 Route Coordinate Matcher")
st.caption("โยนไฟล์ route → ได้พิกัดกลับมาทันที")

if not MASTER_SHEET_CSV_URL:
    st.error("⚠️ ยังไม่ได้ตั้งค่า MASTER_SHEET_CSV_URL — ดูวิธีตั้งค่าใน DEPLOY_INSTRUCTIONS.md")
    st.stop()

try:
    master_df = load_master_data(MASTER_SHEET_CSV_URL)
    st.success(f"✅ เชื่อมต่อ Master Data สำเร็จ: {len(master_df):,} รายชื่อลูกค้า")
except Exception as e:
    st.error(f"โหลด Master Data ไม่สำเร็จ: {e}")
    st.stop()

with st.expander("🔍 ดูตัวอย่าง Master Data"):
    st.dataframe(master_df[['Cust ID', 'Ship To Name', 'Latitude', 'Longitude']].head(20))

st.divider()

uploaded_file = st.file_uploader("ลากไฟล์ route มาวางที่นี่ หรือคลิกเพื่อเลือกไฟล์", type=['xls', 'xlsx'])

if uploaded_file is not None:
    with st.spinner("กำลังประมวลผล..."):
        try:
            sheets, report, unmatched = process_route_file(uploaded_file.read(), master_df)

            total_matched = sum(r['matched'] for r in report)
            total_rows = sum(r['total'] for r in report)
            st.success(f"เสร็จแล้ว! จับคู่พิกัดได้ {total_matched:,} / {total_rows:,} แถว")

            cols = st.columns(len(report))
            for col, r in zip(cols, report):
                with col:
                    if r['skipped']:
                        st.metric(f"Sheet: {r['sheet']}", "ข้าม", "ไม่พบคอลัมน์ Cust Code")
                    else:
                        st.metric(f"Sheet: {r['sheet']}", f"{r['matched']}/{r['total']}",
                                   f"ShipTo:{r['via_ship']} Code:{r['via_code']}")

            output_bytes = to_excel_bytes(sheets)
            base_name = uploaded_file.name.rsplit('.', 1)[0]
            st.download_button(
                "⬇️ ดาวน์โหลดไฟล์พร้อมพิกัด",
                data=output_bytes,
                file_name=f"{base_name}_with_coordinates.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            if unmatched:
                st.warning(f"⚠️ มี {len(unmatched)} รายการที่ยังไม่มีพิกัด")
                unmatched_df = pd.DataFrame(unmatched).drop_duplicates(subset=['Cust Code'])
                st.dataframe(unmatched_df, use_container_width=True)
                st.info("💡 นำ Cust Code เหล่านี้ไปเพิ่มพิกัดใน Google Sheets Master Data แล้วรีเฟรชหน้านี้")

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            st.exception(e)

st.divider()
st.caption(f"Master Data อัพเดตล่าสุด (cache): {datetime.now().strftime('%Y-%m-%d %H:%M')} — ข้อมูลรีเฟรชทุก 5 นาที")
