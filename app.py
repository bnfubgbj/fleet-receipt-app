"""
ระบบบันทึกการซ่อมบำรุง Fleet รถบริษัท - แบบ "โรงรถ" คลิกเลือกรถแล้วลากไฟล์เข้า
=============================================================
หน้าแรก: การ์ดรถทั้งหมด คลิกคันที่ต้องการ / เพิ่มรถใหม่ / ทำเครื่องหมายขายแล้ว
หน้าอัปโหลด: ลากรูป/PDF ใบเสร็จเข้ามา -> AI อ่านข้อมูลให้ -> เช็ค/แก้ -> บันทึก

วิธีติดตั้ง: ดู README.md
"""

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import google.generativeai as genai
import json
import io
from datetime import date, datetime, timedelta

# ---------------- ตั้งค่าพื้นฐาน ----------------
st.set_page_config(page_title="โรงรถ - บันทึกซ่อมบำรุง", page_icon="🚚", layout="wide")

SHEET_VEHICLES = "Vehicles"
SHEET_LOG = "Maintenance_Log"
SHEET_PM = "PM_Schedule"

MAINT_TYPES = ["PM ตามรอบ", "ซ่อมฉุกเฉิน", "เปลี่ยนอะไหล่", "อื่นๆ"]
VEHICLE_TYPES = ["6 ล้อ", "กระบะ"]
DEPARTMENTS = ["ฝ่ายขาย", "ฝ่ายจัดส่ง", "อื่นๆ"]
DRIVE_FOLDER_NAME = "Fleet Maintenance Receipts"

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
  "plate_hint": "เลขทะเบียนรถ ถ้ามีระบุในใบเสร็จ (null ถ้าไม่มี)"
}

ถ้าใบเสร็จแยกค่าแรง/ค่าอะไหล่ไม่ชัดเจน ให้ใส่ยอดรวมทั้งหมดใน "parts" และ "labor" เป็น 0
ตอบเป็น JSON ล้วนๆ เท่านั้น ไม่ต้องมี ```json หรือคำอธิบายใดๆ"""


# ---------------- เชื่อมต่อ Google Sheets + Drive ----------------
@st.cache_resource
def get_credentials():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    return Credentials.from_service_account_info(creds_dict, scopes=scopes)


@st.cache_resource
def get_sheet_client():
    return gspread.authorize(get_credentials())


@st.cache_resource
def get_spreadsheet():
    client = get_sheet_client()
    return client.open_by_key(st.secrets["SPREADSHEET_ID"])


@st.cache_resource
def get_drive_service():
    return build("drive", "v3", credentials=get_credentials())


@st.cache_resource
def get_drive_folder_id():
    service = get_drive_service()
    query = (
        f"name = '{DRIVE_FOLDER_NAME}' and "
        "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    )
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]
    folder_metadata = {"name": DRIVE_FOLDER_NAME, "mimeType": "application/vnd.google-apps.folder"}
    folder = service.files().create(body=folder_metadata, fields="id").execute()
    return folder["id"]


def upload_receipt_to_drive(file_bytes, mime_type, plate, record_date, shop):
    service = get_drive_service()
    folder_id = get_drive_folder_id()
    ext = "pdf" if mime_type == "application/pdf" else mime_type.split("/")[-1]
    safe_shop = (shop or "ไม่ระบุอู่").replace("/", "-")[:40]
    filename = f"{plate}_{record_date}_{safe_shop}.{ext}"
    file_metadata = {"name": filename, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
    uploaded = service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()
    service.permissions().create(fileId=uploaded["id"], body={"role": "reader", "type": "anyone"}).execute()
    return uploaded.get("webViewLink", "")


DEMO_VEHICLES = [
    {"plate": "ขาย-6ล้อ-01", "type": "6 ล้อ", "department": "ฝ่ายขาย", "status": "ใช้งาน"},
    {"plate": "ขาย-6ล้อ-02", "type": "6 ล้อ", "department": "ฝ่ายขาย", "status": "ใช้งาน"},
    {"plate": "จัดส่ง-กะบะ-01", "type": "กระบะ", "department": "ฝ่ายจัดส่ง", "status": "ใช้งาน"},
    {"plate": "จัดส่ง-6ล้อ-01", "type": "6 ล้อ", "department": "ฝ่ายจัดส่ง", "status": "ใช้งาน"},
]


def get_vehicles_full(include_sold=False):
    """คืนค่า (รายชื่อรถ [{plate,type,department,status}], เชื่อมต่อ Sheet สำเร็จหรือไม่)"""
    try:
        sh = get_spreadsheet()
        ws = sh.worksheet(SHEET_VEHICLES)
        rows = ws.get_all_records()
        vehicles = []
        for r in rows:
            plate = r.get("ทะเบียนรถ")
            if not plate:
                continue
            status = r.get("สถานะ", "ใช้งาน") or "ใช้งาน"
            if not include_sold and status == "ขายแล้ว":
                continue
            vehicles.append({
                "plate": plate,
                "type": r.get("ประเภท", ""),
                "department": r.get("แผนก", ""),
                "status": status,
            })
        return vehicles, True
    except Exception:
        return DEMO_VEHICLES, False


def add_vehicle_to_sheet(plate, vtype, department):
    sh = get_spreadsheet()
    ws = sh.worksheet(SHEET_VEHICLES)
    ws.append_row([plate, vtype, department, "", 0, "ใช้งาน"])


def mark_vehicle_sold(plate):
    sh = get_spreadsheet()
    ws = sh.worksheet(SHEET_VEHICLES)
    cell = ws.find(plate)
    if cell:
        ws.update_cell(cell.row, 6, "ขายแล้ว")  # คอลัมน์ F = สถานะ


def update_vehicle(old_plate, new_plate, new_type, new_department):
    sh = get_spreadsheet()
    vws = sh.worksheet(SHEET_VEHICLES)
    cell = vws.find(old_plate)
    if not cell:
        raise ValueError("ไม่พบรถคันนี้ในระบบ")
    vws.update_cell(cell.row, 1, new_plate)
    vws.update_cell(cell.row, 2, new_type)
    vws.update_cell(cell.row, 3, new_department)

    if new_plate != old_plate:
        # เปลี่ยนทะเบียน -> อัปเดตให้ตรงกันในชีตอื่นด้วย
        pws = sh.worksheet(SHEET_PM)
        pcell = pws.find(old_plate)
        if pcell:
            pws.update_cell(pcell.row, 1, new_plate)

        lws = sh.worksheet(SHEET_LOG)
        try:
            matches = lws.findall(old_plate)
        except Exception:
            matches = []
        for m in matches:
            if m.col == 2:  # คอลัมน์ B = ทะเบียนรถ ใน Maintenance_Log
                lws.update_cell(m.row, 2, new_plate)


def append_maintenance_record(record):
    sh = get_spreadsheet()
    ws = sh.worksheet(SHEET_LOG)
    ws.append_row([
        record["date"], record["plate"], record["type"], record["description"],
        record["labor"], record["parts"], record["total"], record["shop"],
        record["mileage"] or "", record["downtime_days"], record["recorded_by"],
        record.get("receipt_link", "")
    ])
    if record["mileage"]:
        vws = sh.worksheet(SHEET_VEHICLES)
        cell = vws.find(record["plate"])
        if cell:
            vws.update_cell(cell.row, 5, record["mileage"])
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
    pws.update_cell(row, 4, record_date)
    pws.update_cell(row, 5, mileage or "")
    pws.update_cell(row, 6, next_date.strftime("%Y-%m-%d"))
    days_left = (next_date - datetime.now()).days
    status = "ปกติ"
    if days_left < 0:
        status = "เลยกำหนด!"
    elif days_left <= 7:
        status = "ใกล้ครบรอบ"
    pws.update_cell(row, 7, status)


def extract_receipt_data(file_bytes, mime_type):
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model = genai.GenerativeModel("gemini-flash-latest")
    file_part = {"mime_type": mime_type, "data": file_bytes}
    response = model.generate_content(
        [file_part, EXTRACTION_PROMPT],
        generation_config={"response_mime_type": "application/json"}
    )
    text = response.text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(text)


def vehicle_icon(vtype):
    return "🚛" if "6" in (vtype or "") else "🛻"


# ================= Session State =================
for key, default in [
    ("view", "garage"), ("selected_plate", None), ("extracted", None),
    ("file_bytes", None), ("mime_type", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


def go_to_upload(plate):
    st.session_state.selected_plate = plate
    st.session_state.view = "upload"
    st.session_state.extracted = None
    st.session_state.file_bytes = None
    st.rerun()


def go_to_garage():
    st.session_state.view = "garage"
    st.session_state.selected_plate = None
    st.session_state.extracted = None
    st.session_state.file_bytes = None
    st.rerun()


def go_to_add_vehicle():
    st.session_state.view = "add_vehicle"
    st.rerun()


# ================= หน้าโรงรถ =================
def render_garage():
    st.title("🚚 โรงรถ — เลือกคันที่ต้องการบันทึกซ่อมบำรุง")
    vehicles, sheet_connected = get_vehicles_full()

    if not sheet_connected:
        st.warning(
            "🔌 โหมดทดลอง: ยังไม่ได้เชื่อมต่อ Google Sheet — แสดงรถตัวอย่างให้ลองระบบก่อน "
            "การเพิ่ม/ขายรถจะใช้ไม่ได้จนกว่าจะเชื่อม Sheet จริง (ดู README.md)"
        )

    st.write("")
    all_vehicles_incl_sold, _ = get_vehicles_full(include_sold=True)
    existing_plates_all = {veh["plate"] for veh in all_vehicles_incl_sold}
    cols_per_row = 3
    items = list(vehicles) + [{"plate": None}]  # การ์ดสุดท้าย = เพิ่มรถใหม่

    for i in range(0, len(items), cols_per_row):
        row_items = items[i:i + cols_per_row]
        cols = st.columns(cols_per_row)
        for col, v in zip(cols, row_items):
            with col:
                if v["plate"] is None:
                    with st.container(border=True):
                        st.markdown(
                            "<div style='text-align:center; font-size:72px; line-height:1.1; color:#888'>➕</div>",
                            unsafe_allow_html=True
                        )
                        st.markdown(
                            "<h3 style='text-align:center; margin-bottom:0; color:#888'>เพิ่มรถใหม่</h3>",
                            unsafe_allow_html=True
                        )
                        st.markdown("<p style='text-align:center; color:gray;'>&nbsp;</p>", unsafe_allow_html=True)
                        if st.button("➕ เพิ่มรถ", key="add_vehicle_btn", use_container_width=True,
                                     disabled=not sheet_connected):
                            go_to_add_vehicle()
                    continue

                with st.container(border=True):
                    st.markdown(
                        f"<div style='text-align:center; font-size:72px; line-height:1.1'>{vehicle_icon(v['type'])}</div>",
                        unsafe_allow_html=True
                    )
                    st.markdown(
                        f"<h3 style='text-align:center; margin-bottom:0'>{v['plate']}</h3>",
                        unsafe_allow_html=True
                    )
                    st.markdown(
                        f"<p style='text-align:center; color:gray; margin-top:0'>{v['department']} • {v['type']}</p>",
                        unsafe_allow_html=True
                    )
                    if st.button("📤 อัปโหลดใบเสร็จ", key=f"select_{v['plate']}", use_container_width=True):
                        go_to_upload(v["plate"])

                    with st.expander("✏️ แก้ไขข้อมูลรถ"):
                        st.caption("ใช้ตอนเปลี่ยนจากชื่อชั่วคราวเป็นทะเบียนจริง")
                        new_plate = st.text_input(
                            "ทะเบียนรถ", value=v["plate"], key=f"edit_plate_{v['plate']}"
                        )
                        new_type = st.radio(
                            "ประเภท", options=VEHICLE_TYPES,
                            index=VEHICLE_TYPES.index(v["type"]) if v["type"] in VEHICLE_TYPES else 0,
                            format_func=lambda t: f"{vehicle_icon(t)}  {t}",
                            horizontal=True, key=f"edit_type_{v['plate']}"
                        )
                        dept_options = DEPARTMENTS if v["department"] in DEPARTMENTS else DEPARTMENTS + [v["department"]]
                        new_department = st.radio(
                            "แผนก", options=dept_options,
                            index=dept_options.index(v["department"]) if v["department"] in dept_options else 0,
                            horizontal=True, key=f"edit_dept_{v['plate']}"
                        )
                        if st.button("💾 บันทึกการแก้ไข", key=f"save_edit_{v['plate']}",
                                      use_container_width=True, disabled=not sheet_connected):
                            new_plate_clean = new_plate.strip()
                            if not new_plate_clean:
                                st.warning("กรุณากรอกทะเบียนรถ")
                            elif new_plate_clean != v["plate"] and new_plate_clean in existing_plates_all:
                                st.warning(f"มีทะเบียน {new_plate_clean} อยู่ในระบบแล้ว")
                            else:
                                try:
                                    update_vehicle(v["plate"], new_plate_clean, new_type, new_department)
                                    st.success("บันทึกการแก้ไขเรียบร้อยแล้ว!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"แก้ไขไม่สำเร็จ: {e}")

                    with st.expander("🚫 รถคันนี้ขายไปแล้ว?"):
                        st.caption("รถจะถูกซ่อนออกจากหน้านี้ (ประวัติซ่อมเก่ายังอยู่ครบใน Sheet)")
                        if st.button("ยืนยัน — ทำเครื่องหมายว่าขายแล้ว", key=f"sold_{v['plate']}",
                                      use_container_width=True, disabled=not sheet_connected):
                            try:
                                mark_vehicle_sold(v["plate"])
                                st.success(f"ทำเครื่องหมาย {v['plate']} ว่าขายแล้ว")
                                st.rerun()
                            except Exception as e:
                                st.error(f"ทำรายการไม่สำเร็จ: {e}")


# ================= หน้าเพิ่มรถใหม่ =================
def render_add_vehicle():
    if st.button("← กลับไปโรงรถ"):
        go_to_garage()

    st.title("➕ เพิ่มรถคันใหม่")

    vehicles, sheet_connected = get_vehicles_full(include_sold=True)
    existing_plates = {v["plate"] for v in vehicles}

    vtype_label = st.radio(
        "เลือกไอคอนรถ",
        options=VEHICLE_TYPES,
        format_func=lambda t: f"{vehicle_icon(t)}  {t}",
        horizontal=True,
    )

    plate = st.text_input("ทะเบียนรถ")

    department = st.radio("แผนก", options=DEPARTMENTS, horizontal=True)
    custom_department = ""
    if department == "อื่นๆ":
        custom_department = st.text_input("ระบุชื่อแผนก")

    final_department = custom_department if department == "อื่นๆ" and custom_department else department

    if st.button("✅ เพิ่มรถคันนี้", type="primary", use_container_width=True, disabled=not sheet_connected):
        if not plate.strip():
            st.warning("กรุณากรอกทะเบียนรถ")
        elif plate.strip() in existing_plates:
            st.warning(f"มีทะเบียน {plate} อยู่ในระบบแล้ว กรุณาตรวจสอบ")
        else:
            try:
                add_vehicle_to_sheet(plate.strip(), vtype_label, final_department)
                st.success(f"เพิ่มรถ {plate} เรียบร้อยแล้ว!")
                go_to_garage()
            except Exception as e:
                st.error(f"เพิ่มรถไม่สำเร็จ: {e}")

    if not sheet_connected:
        st.caption("ปุ่มเพิ่มรถถูกปิดไว้ชั่วคราว เพราะยังไม่ได้เชื่อมต่อ Google Sheet")


# ================= หน้าอัปโหลด =================
def render_upload():
    plate = st.session_state.selected_plate
    _, sheet_connected = get_vehicles_full()

    if st.button("← กลับไปเลือกรถ"):
        go_to_garage()

    st.title(f"🔧 บันทึกซ่อมบำรุง: {plate}")
    st.caption("ลากรูปหรือไฟล์ PDF ใบเสร็จเข้ามา ระบบจะอ่านข้อมูลให้อัตโนมัติ")

    if not sheet_connected:
        st.warning(
            "🔌 โหมดทดลอง: ยังไม่ได้เชื่อมต่อ Google Sheet — ลองดูว่า AI อ่านข้อมูลถูกต้องแค่ไหนได้ "
            "แต่กด **บันทึก** ไม่ได้จนกว่าจะตั้งค่า Google Sheet ให้ครบ (ดู README.md)"
        )

    uploaded_file = st.file_uploader(
        "ลากไฟล์มาวาง หรือกดเพื่อถ่ายรูป",
        type=["jpg", "jpeg", "png", "webp", "pdf"],
        accept_multiple_files=False,
        key=f"uploader_{plate}"
    )

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        mime_type = uploaded_file.type
        is_pdf = mime_type == "application/pdf"

        if st.session_state.file_bytes != file_bytes:
            st.session_state.file_bytes = file_bytes
            st.session_state.mime_type = mime_type
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

        plate_hint = data.get("plate_hint")
        if plate_hint and plate_hint.replace(" ", "") not in plate.replace(" ", ""):
            st.warning(f"⚠️ ใบเสร็จนี้ระบุทะเบียนว่า **{plate_hint}** แต่กำลังบันทึกให้คันที่เลือกคือ **{plate}** เช็คให้แน่ใจว่าเลือกถูกคัน")

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
            if not recorded_by:
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
                    with st.spinner("กำลังอัปโหลดใบเสร็จต้นฉบับเก็บไว้..."):
                        try:
                            record["receipt_link"] = upload_receipt_to_drive(
                                st.session_state.file_bytes, st.session_state.mime_type,
                                plate, record["date"], shop
                            )
                        except Exception as drive_err:
                            record["receipt_link"] = ""
                            st.warning(f"บันทึกข้อมูลได้ แต่เก็บไฟล์ต้นฉบับไม่สำเร็จ: {drive_err}")
                    append_maintenance_record(record)
                    st.success(f"บันทึกข้อมูลรถ {plate} เรียบร้อยแล้ว!")
                    st.balloons()
                    st.session_state.extracted = None
                    st.session_state.file_bytes = None
                except Exception as e:
                    st.error(f"บันทึกไม่สำเร็จ: {e}")

        if not sheet_connected:
            st.caption("ปุ่มบันทึกถูกปิดไว้ชั่วคราว เพราะยังไม่ได้เชื่อมต่อ Google Sheet")


# ================= Router =================
if st.session_state.view == "garage":
    render_garage()
elif st.session_state.view == "add_vehicle":
    render_add_vehicle()
else:
    render_upload()

st.divider()
st.caption("ระบบบันทึกการซ่อมบำรุง Fleet | ข้อมูลจะถูกบันทึกเข้า Google Sheet โดยตรง")
