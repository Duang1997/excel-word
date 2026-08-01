import streamlit as st
import pandas as pd
import re
import io
import dataframe_image as dfi
from PIL import Image
import pytesseract
from streamlit_paste_button import paste_image_button
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt

# ==========================================
# 0. ตั้งค่าฟอนต์ภาษาไทยสำหรับการส่งออกภาพ (ต้องมีไฟล์ THSarabunNew.ttf)
# ==========================================
try:
    font_path = "THSarabunNew.ttf"
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.family'] = fm.FontProperties(fname=font_path).get_name()
except Exception:
    pass # หากไม่มีไฟล์ฟอนต์ ระบบจะทำงานต่อได้แต่อาจแสดงผลภาษาไทยผิดพลาด

# ==========================================
# 1. ฟังก์ชันจัดรูปแบบข้อมูลสำหรับแสดงผลบน Word
# ==========================================
def format_thai_date(date_str):
    thai_months = [
        "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
        "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
    ]
    try:
        if isinstance(date_str, str) and '/' in date_str:
            day, month, year = date_str.split('/')
            thai_year = int(year) + 543
            return f"{int(day)} {thai_months[int(month)]} {thai_year}"
        return str(date_str)
    except:
        return str(date_str)

def format_thai_bank(bank_abbr):
    bank_mapping = {
        "KTB": "กรุงไทย",
        "KBANK": "กสิกรไทย",
        "SCB": "ไทยพาณิชย์",
        "BBL": "กรุงเทพ",
        "GSB": "ออมสิน"
    }
    return bank_mapping.get(str(bank_abbr).upper(), str(bank_abbr))

# ==========================================
# 2. ฟังก์ชันจัดรูปแบบข้อมูลสำหรับฐานข้อมูล/ตาราง (มาตรฐานการสืบสวน)
# ==========================================
def clean_date_to_ce(date_val):
    """แปลงวันที่เป็น ค.ศ. (dd/mm/yyyy)"""
    if not isinstance(date_val, str):
        date_val = str(date_val)
    match = re.search(r'(\d{2}/\d{2}/)(\d{4})', date_val)
    if match:
        day_month = match.group(1)
        year = int(match.group(2))
        if year > 2400: 
            year -= 543
        return f"{day_month}{year}"
    return date_val

def standardize_bank_name(bank_name):
    """แปลงชื่อธนาคารเป็นตัวย่อภาษาอังกฤษ"""
    bank_mapping = {
        "ธนาคารกสิกรไทย": "KBANK", "กสิกรไทย": "KBANK",
        "ธนาคารไทยพาณิชย์": "SCB", "ไทยพาณิชย์": "SCB",
        "ธนาคารกรุงเทพ": "BBL", "กรุงเทพ": "BBL",
        "ธนาคารกรุงไทย": "KTB", "กรุงไทย": "KTB",
        "ธนาคารออมสิน": "GSB", "ออมสิน": "GSB"
    }
    return bank_mapping.get(str(bank_name), str(bank_name))

def style_statement_table(df):
    """กำหนดรูปแบบตารางให้คล้าย Excel (มีเส้นขอบ, จัดกึ่งกลาง)"""
    styles = [
        dict(selector="th", props=[("border", "1px solid black"), ("text-align", "center"), ("background-color", "#ffffff")]),
        dict(selector="td", props=[("border", "1px solid black"), ("text-align", "center")])
    ]
    return df.style.set_table_styles(styles).hide(axis="index")

# ==========================================
# 3. ฟังก์ชันประมวลผลข้อมูลหลัก
# ==========================================
def process_statement_data(raw_text):
    records = []
    lines = raw_text.strip().split('\n')
    
    for line in lines:
        if not line.strip(): 
            continue
            
        parts = re.split(r'\t|\s{2,}', line.strip())
        if len(parts) >= 11:
            record = {
                "วันที่ทำรายการ": clean_date_to_ce(parts[0]),
                "เวลาที่ทำรายการ": parts[1],
                "ประเภทรายการ": parts[2],
                "ช่องทาง": parts[3],
                "ชื่อธนาคารต้นทาง": standardize_bank_name(parts[4]),
                "หมายเลขบัญชีต้นทาง": str(parts[5]), # บังคับเป็น String
                "ชื่อบัญชีต้นทาง": parts[6],
                "ชื่อธนาคารปลายทาง": standardize_bank_name(parts[7]),
                "หมายเลขบัญชีปลายทาง": str(parts[8]), # บังคับเป็น String
                "ชื่อบัญชีปลายทาง": parts[9],
                "ยอดเงิน": parts[10],
                "คงเหลือ": parts[11] if len(parts) >= 12 else "-"
            }
            records.append(record)
    return pd.DataFrame(records)

def process_excel_upload(df_raw):
    records = []
    df_raw = df_raw.astype(str)
    
    for _, row in df_raw.iterrows():
        cols = row.tolist()
        if len(cols) >= 11:
            record = {
                "วันที่ทำรายการ": clean_date_to_ce(cols[0]),
                "เวลาที่ทำรายการ": cols[1],
                "ประเภทรายการ": cols[2],
                "ช่องทาง": cols[3],
                "ชื่อธนาคารต้นทาง": standardize_bank_name(cols[4]),
                "หมายเลขบัญชีต้นทาง": str(cols[5]).replace('.0', ''), 
                "ชื่อบัญชีต้นทาง": cols[6],
                "ชื่อธนาคารปลายทาง": standardize_bank_name(cols[7]),
                "หมายเลขบัญชีปลายทาง": str(cols[8]).replace('.0', ''),
                "ชื่อบัญชีปลายทาง": cols[9],
                "ยอดเงิน": cols[10],
                "คงเหลือ": cols[11] if len(cols) >= 12 else "-"
            }
            records.append(record)
    return pd.DataFrame(records)

# ==========================================
# 4. ส่วนติดต่อผู้ใช้งาน (UI)
# ==========================================
st.set_page_config(page_title="ระบบวิเคราะห์รายการเดินบัญชี", layout="wide")
st.title("ระบบวิเคราะห์และสกัดข้อมูลรายการเดินบัญชี")

st.subheader("⚙️ กำหนดรูปแบบข้อความ (Template)")
default_template = "เมื่อวันที่ {date} เวลาประมาณ {time} น. บัญชีธนาคาร{src_bank} หมายเลขบัญชี {src_acc} ชื่อบัญชี {src_name} ได้ทำการโอนเงินไปยัง บัญชีธนาคาร{dst_bank} หมายเลขบัญชี {dst_acc} ชื่อบัญชี {dst_name} จำนวน {amount} บาท"
user_template = st.text_area("ตัวแปรที่รองรับ: {date}, {time}, {type}, {channel}, {src_bank}, {src_acc}, {src_name}, {dst_bank}, {dst_acc}, {dst_name}, {amount}, {balance}", value=default_template)

st.subheader("📥 นำเข้าข้อมูล")
tab1, tab2, tab3 = st.tabs(["1. วางข้อความ", "2. อัปโหลด/ลากไฟล์ Excel", "3. วางภาพจาก Clipboard"])

df_result = pd.DataFrame()

with tab1:
    raw_text = st.text_area("วางข้อมูลรายการเดินบัญชี (ข้อความ/ตาราง) ลงในช่องนี้", height=200)
    if st.button("ประมวลผลจากข้อความ"):
        if raw_text:
            df_result = process_statement_data(raw_text)

with tab2:
    uploaded_excel = st.file_uploader("ลากไฟล์ Excel (.xlsx, .xls) มาวางที่นี่", type=['xlsx', 'xls'])
    if st.button("ประมวลผลจาก Excel"):
        if uploaded_excel is not None:
            df_raw = pd.read_excel(uploaded_excel, dtype=str)
            df_result = process_excel_upload(df_raw)
        else:
            st.warning("กรุณาอัปโหลดไฟล์ Excel ก่อนดำเนินการ")

with tab3:
    st.info("คัดลอกรูปภาพ (Ctrl+C) จากนั้นคลิกที่ปุ่มด้านล่างเพื่อดึงภาพจากคลิปบอร์ด")
    paste_result = paste_image_button(
        label="📋 คลิกเพื่อวางภาพจาก Clipboard",
        text_color="#ffffff",
        background_color="#28a745",
        hover_background_color="#218838"
    )
    
    if paste_result.image_data is not None:
        image = paste_result.image_data
        st.image(image, caption="ภาพที่นำเข้า", use_column_width=True)
        
        if st.button("ประมวลผลจากภาพ (OCR)"):
            try:
                extracted_text = pytesseract.image_to_string(image, lang='tha+eng')
                df_result = process_statement_data(extracted_text)
            except Exception as e:
                st.error("เกิดข้อผิดพลาดในการอ่านภาพ โปรดตรวจสอบการติดตั้ง Tesseract-OCR")

# ==========================================
# 5. ส่วนแสดงผลและส่งออก
# ==========================================
if not df_result.empty:
    st.success("ประมวลผลสำเร็จ")
    st.divider()
    
    st.subheader("📝 ข้อความสำหรับรายงานสืบสวน (Word)")
    generated_text = ""
    for _, row in df_result.iterrows():
        try:
            row_text = user_template.format(
                date=format_thai_date(row['วันที่ทำรายการ']),
                time=row['เวลาที่ทำรายการ'],
                type=row['ประเภทรายการ'],
                channel=row['ช่องทาง'],
                src_bank=format_thai_bank(row['ชื่อธนาคารต้นทาง']),
                src_acc=row['หมายเลขบัญชีต้นทาง'],
                src_name=row['ชื่อบัญชีต้นทาง'],
                dst_bank=format_thai_bank(row['ชื่อธนาคารปลายทาง']),
                dst_acc=row['หมายเลขบัญชีปลายทาง'],
                dst_name=row['ชื่อบัญชีปลายทาง'],
                amount=row['ยอดเงิน'],
                balance=row['คงเหลือ']
            )
            generated_text += row_text + "\n"
        except Exception as e:
            st.error(f"รูปแบบ Template ไม่ถูกต้อง: {e}")
            break
    
    if generated_text:
        st.text_area("คัดลอกข้อความด้านล่าง:", value=generated_text, height=200)
    
    st.divider()
    
    st.subheader("📊 ข้อมูลตาราง")
    st.dataframe(df_result)
    
    col1, col2 = st.columns(2)
    
    with col1:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='Statement')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=excel_buffer.getvalue(),
            file_name="Statement_Processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col2:
        image_buffer = io.BytesIO()
        try:
            styled_df = style_statement_table(df_result)
            dfi.export(styled_df, image_buffer, table_conversion="matplotlib")
            st.download_button(
                label="🖼️ ดาวน์โหลดรูปภาพตาราง",
                data=image_buffer.getvalue(),
                file_name="Statement_Table.png",
                mime="image/png"
            )
        except Exception as e:
            st.warning(f"ไม่สามารถส่งออกภาพได้: {e} โปรดตรวจสอบไฟล์ฟอนต์หรือไลบรารี")
