"""
ระบบบันทึกการซ่อมบำรุง Fleet รถบริษัท - แบบถ่ายรูปใบเสร็จ
=============================================================
ลากรูปใบเสร็จเข้ามา -> AI อ่านข้อมูลให้ -> เช็ค/แก้ -> บันทึกเข้า Google Sheet

วิธีติดตั้ง: ดู README.md
"""

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import google.generativeai as genai
import base64
import json
from datetime import date, datetime, timedelta
from PIL import Image
import io

# ---------------- ตั้งค่าพื้นฐาน ----------------
st.set_page_config(page_title="บันทึกซ่อมบำรุงรถ", page_icon="🔧", layout="centered")

SHEET_VEHICLES = "Vehicles"
SHEET_LOG = "Maintenance_Log"
SHEET_PM = "PM_Schedule"

MAINT_TYPES = ["PM ตามรอบ", "ซ่อมฉุกเฉิน", "เปลี่ยนอะไหล่", "อื่นๆ"]

EXTRACTION_PROMPT = """คุณคือผู้ช่วยอ่านใบเสร็จ/ใบแจ้งหนี้ซ่อมรถยนต์ภาษาไทย
อ่านรูปใบเสร็จนี้แล้วดึงข้อมูลออกมาเป็น JSON เท่านั้น ห้ามมีข้อความอื่นนอกเหนือจาก JSON

สำคัญมาก: ใบเสร็จไทยส่วนใหญ่ใช้ปีพุทธศักราช (พ.ศ.) เช่น "11 มิถุนายน 2569"
ต้องแปลงเป็นปีคริสต์ศักราช (ค.ศ.) เสมอ โดยเอาปี พ.ศ. ลบด้วย 543 ก่อนใส่ในฟิลด์ "date"
ตัวอย่าง: "11 มิถุนายน 2569" (พ.ศ.) = 11 มิถุนายน ค.ศ. 2026 -> ใส่เป็น "2026-06-11"
ถ้าใบเสร็จระบุปีเป็น ค.ศ. อยู่แล้ว (เช่น 2026) ไม่ต้องแปลงซ้ำ

รูปแบบ JSON ที่ต้องการ:
{
  "date": "YYYY-MM-DD (ค.ศ. เท่านั้น แปลงจาก พ.ศ. แล้ว) หรือ null ถ้าไม่พบ",
  "shop": "ชื่ออู่/ศูนย์บริการ หรือ null",
  "description": "สรุปรายการซ่อม/เปลี่ยนอะไหล่ทั้งหมดเป็นข้อความสั้นๆ ภาษาไทย",
  "labor": ตัวเลขค่าแรง (0 ถ้าไม่พบหรือรวมกับค่าอะไหล่),
  "parts": ตัวเลขค่าอะไหล่ (0 ถ้าไม่พบ),
  "total": ตัวเลขยอดรวมทั้งหมดที่ระบุในใบเสร็จ (รวม VAT ถ้ามี),
  "mileage": ตัวเลขไมล์/เลขกิโลเมตร ถ้ามีระบุในใบเสร็จ (null ถ้าไม่มี),
  "plate_hint": "เลขทะเบียนรถ ถ้ามีระบุในใบเสร็จ (null ถ้าไม่มี) ใช้แค่ช่วยเทียบ ไม่ใช้แทนการเลือกทะเบียนเอง"
}

ถ้าใบเสร็จแยกค่าแรง/ค่าอะไหล่ไม่ชัดเจน ให้ใส่ยอดรวมทั้งหมดใน "parts" และ "labor" เป็น 0
ตอบเป็น JSON ล้วนๆ เท่านั้น ไม่ต้องมี ```json หรือคำอธิบายใดๆ"""


# ---------------- เชื่อมต่อ Google Sheets ----------------
@st.cache_resource
def get_sheet_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


@st.cache_resource
def get_spreadsheet():
    client = get_sheet_client()
    return client.open_by_key(st.secrets["SPREADSHEET_ID"])


def get_vehicle_list():
    """คืนค่า (รายชื่อทะเบียนรถ, เชื่อมต่อ Sheet สำเร็จหรือไม่)"""
    try:
        sh = get_spreadsheet()
        ws = sh.worksheet(SHEET_VEHICLES)
        rows = ws.get_all_records()
        return [r["ทะเบียนรถ"] for r in rows if r.get("ทะเบียนรถ")], True
    except Exception:
        return [], False


def get_vehicle_department(plate):
    sh = get_spreadsheet()
    ws = sh.worksheet(SHEET_VEHICLES)
    rows = ws.get_all_records()
    for r in rows:
        if r.get("ทะเบียนรถ") == plate:
            return r.get("แผนก", "")
    return ""


def append_maintenance_record(record):
    sh = get_spreadsheet()
    ws = sh.worksheet(SHEET_LOG)
    ws.append_row([
        record["date"], record["plate"], record["type"], record["description"],
        record["labor"], record["parts"], record["total"], record["shop"],
        record["mileage"] or "", record["downtime_days"], record["recorded_by"]
    ])

    # อัปเดตเลขไมล์ล่าสุดในแท็บ Vehicles
    if record["mileage"]:
        vws = sh.worksheet(SHEET_VEHICLES)
        cell = vws.find(record["plate"])
        if cell:
            vws.update_cell(cell.row, 5, record["mileage"])

    # ถ้าเป็นงาน PM ตามรอบ อัปเดต PM_Schedule ให้ด้วย
    if record["type"] == "PM ตามรอบ":
        update_pm_schedule(record["plate"], record["date"], record["mileage"])


def update_pm_schedule(plate, record_date, mileage):
    sh = get_spreadsheet()
    pws = sh.worksheet(SHEET_PM)
    cell = pws.find(plate)
    if not cell:
        return
    row = cell.row
    interval_months = pws.cell(row, 3).value or 3
    try:
        interval_months = int(interval_months)
    except ValueError:
        interval_months = 3

    last_date = datetime.strptime(record_date, "%Y-%m-%d")
    next_date = last_date + timedelta(days=30 * interval_months)

    pws.update_cell(row, 4, record_date)              # ซ่อม PM ล่าสุด
    pws.update_cell(row, 5, mileage or "")             # เลขไมล์ล่าสุดตอน PM
    pws.update_cell(row, 6, next_date.strftime("%Y-%m-%d"))  # กำหนดครั้งถัดไป

    days_left = (next_date - datetime.now()).days
    status = "ปกติ"
    if days_left < 0:
        status = "เลยกำหนด!"
    elif days_left <= 7:
        status = "ใกล้ครบรอบ"
    pws.update_cell(row, 7, status)


# ---------------- อ่านใบเสร็จด้วย AI (Gemini free tier) — รองรับทั้งรูปภาพและ PDF ----------------
def extract_receipt_data(file_bytes, mime_type):
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model = genai.GenerativeModel("gemini-2.5-flash")

    # ส่งไฟล์แบบ blob ตรงๆ ใช้ได้ทั้งรูปภาพ (jpg/png/webp) และ PDF
    file_part = {"mime_type": mime_type, "data": file_bytes}

    response = model.generate_content(
        [file_part, EXTRACTION_PROMPT],
        generation_config={"response_mime_type": "application/json"}
    )
    text = response.text.strip()
    text = text.replace("```json", "").replace("```", "").strip()
    return json.loads(text)


# ================= UI =================
st.title("🔧 บันทึกการซ่อมบำรุงรถ")
st.caption("ถ่ายรูปหรือลากใบเสร็จเข้ามา ระบบจะอ่านข้อมูลให้อัตโนมัติ")

if "extracted" not in st.session_state:
    st.session_state.extracted = None
if "image_bytes" not in st.session_state:
    st.session_state.image_bytes = None

uploaded_file = st.file_uploader(
    "ลากรูปใบเสร็จหรือไฟล์ PDF มาวาง หรือกดเพื่อถ่ายรูป",
    type=["jpg", "jpeg", "png", "webp", "pdf"],
    accept_multiple_files=False
)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    mime_type = uploaded_file.type
    is_pdf = mime_type == "application/pdf"

    if st.session_state.image_bytes != file_bytes:
        st.session_state.image_bytes = file_bytes
        if is_pdf:
            st.info(f"📄 ไฟล์ PDF: {uploaded_file.name} ({len(file_bytes)/1024:.0f} KB)")
        else:
            st.image(file_bytes, caption="ใบเสร็จที่อัปโหลด", use_container_width=True)
        with st.spinner("กำลังอ่านข้อมูลจากใบเสร็จ..."):
            try:
                st.session_state.extracted = extract_receipt_data(file_bytes, mime_type)
            except Exception as e:
                st.error(f"อ่านข้อมูลไม่สำเร็จ: {e}")
                st.session_state.extracted = {}
    else:
        if is_pdf:
            st.info(f"📄 ไฟล์ PDF: {uploaded_file.name} ({len(file_bytes)/1024:.0f} KB)")
        else:
            st.image(file_bytes, caption="ใบเสร็จที่อัปโหลด", use_container_width=True)

if st.session_state.extracted is not None:
    st.divider()
    st.subheader("ตรวจสอบข้อมูลก่อนบันทึก")
    st.caption("AI อ่านให้เบื้องต้น กรุณาเช็คความถูกต้องก่อนกดบันทึก")

    data = st.session_state.extracted

    vehicle_list, sheet_connected = get_vehicle_list()

    if not sheet_connected:
        st.warning(
            "🔌 โหมดทดลอง: ยังไม่ได้เชื่อมต่อ Google Sheet — ลองดูว่า AI อ่านข้อมูลถูกต้องแค่ไหนได้ "
            "แต่กด **บันทึก** ไม่ได้จนกว่าจะตั้งค่า Google Sheet ให้ครบ (ดู README.md)"
        )
        plate_hint = data.get("plate_hint")
        if plate_hint:
            st.info(f"📄 ใบเสร็จนี้ระบุทะเบียนรถว่า: **{plate_hint}** (เลือกให้ตรงกันด้านล่าง)")
        plate = st.selectbox(
            "ทะเบียนรถ (ตัวอย่าง — ยังไม่ได้ดึงจาก Sheet จริง)",
            options=["ขาย-6ล้อ-01", "ขาย-6ล้อ-02", "จัดส่ง-กะบะ-01", "จัดส่ง-6ล้อ-01"],
            index=None, placeholder="เลือกทะเบียนรถ"
        )
    else:
        plate_hint = data.get("plate_hint")
        if plate_hint:
            st.info(f"📄 ใบเสร็จนี้ระบุทะเบียนรถว่า: **{plate_hint}** (เลือกให้ตรงกันด้านล่าง)")
        plate = st.selectbox("ทะเบียนรถ", options=vehicle_list, index=None, placeholder="เลือกทะเบียนรถ")

    try:
        parsed_date = datetime.strptime(data.get("date") or "", "%Y-%m-%d").date()
    except ValueError:
        parsed_date = date.today()
    record_date = st.date_input("วันที่", value=parsed_date)

    maint_type = st.selectbox("ประเภทงาน", options=MAINT_TYPES)
    description = st.text_area("รายการที่ทำ", value=data.get("description") or "")

    col1, col2 = st.columns(2)
    with col1:
        labor = st.number_input("ค่าแรง (บาท)", min_value=0, value=int(data.get("labor") or 0))
    with col2:
        parts = st.number_input("ค่าอะไหล่ (บาท)", min_value=0, value=int(data.get("parts") or 0))

    total = labor + parts
    st.info(f"รวมค่าใช้จ่าย: {total:,.0f} บาท")

    shop = st.text_input("อู่/ศูนย์บริการ", value=data.get("shop") or "")
    mileage_val = data.get("mileage")
    mileage = st.number_input("เลขไมล์ ณ วันที่ซ่อม (ถ้ามี)", min_value=0, value=int(mileage_val) if mileage_val else 0)
    downtime_days = st.number_input("วันหยุดวิ่ง (วัน)", min_value=0, value=0)
    recorded_by = st.text_input("ผู้บันทึก")

    if st.button("✅ บันทึกข้อมูล", type="primary", use_container_width=True, disabled=not sheet_connected):
        if not plate:
            st.warning("กรุณาเลือกทะเบียนรถก่อนบันทึก")
        elif not recorded_by:
            st.warning("กรุณากรอกชื่อผู้บันทึก")
        else:
            record = {
                "date": record_date.strftime("%Y-%m-%d"),
                "plate": plate,
                "type": maint_type,
                "description": description,
                "labor": labor,
                "parts": parts,
                "total": total,
                "shop": shop,
                "mileage": mileage if mileage > 0 else None,
                "downtime_days": downtime_days,
                "recorded_by": recorded_by,
            }
            try:
                append_maintenance_record(record)
                st.success(f"บันทึกข้อมูลรถ {plate} เรียบร้อยแล้ว!")
                st.session_state.extracted = None
                st.session_state.image_bytes = None
                st.balloons()
            except Exception as e:
                st.error(f"บันทึกไม่สำเร็จ: {e}")

    if not sheet_connected:
        st.caption("ปุ่มบันทึกถูกปิดไว้ชั่วคราว เพราะยังไม่ได้เชื่อมต่อ Google Sheet")

st.divider()
st.caption("ระบบบันทึกการซ่อมบำรุง Fleet | ข้อมูลจะถูกบันทึกเข้า Google Sheet โดยตรง")
