import streamlit as st
import pandas as pd
import re
import io
import dataframe_image as dfi

# ==========================================
# ฟังก์ชันจัดการข้อมูล
# ==========================================
def format_thai_date(date_str):
    thai_months = [
        "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
        "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
    ]
    try:
        day, month, year = date_str.split('/')
        thai_year = int(year) + 543
        return f"{int(day)} {thai_months[int(month)]} {thai_year}"
    except:
        return date_str

def format_thai_bank(bank_abbr):
    bank_mapping = {
        "KTB": "ธนาคารกรุงไทย",
        "KBANK": "ธนาคารกสิกรไทย",
        "SCB": "ธนาคารไทยพาณิชย์",
        "BBL": "ธนาคารกรุงเทพ",
        "GSB": "ธนาคารออมสิน"
    }
    return bank_mapping.get(bank_abbr, bank_abbr)

def process_statement_data(raw_text):
    records = []
    lines = raw_text.strip().split('\n')
    
    # แปลงชื่อธนาคารให้เป็นตัวย่อมาตรฐานสำหรับฐานข้อมูล
    bank_mapping = {
        "ธนาคารกสิกรไทย": "KBANK", "กสิกรไทย": "KBANK",
        "ธนาคารไทยพาณิชย์": "SCB", "ไทยพาณิชย์": "SCB",
        "ธนาคารกรุงเทพ": "BBL", "กรุงเทพ": "BBL",
        "ธนาคารกรุงไทย": "KTB", "กรุงไทย": "KTB",
        "ธนาคารออมสิน": "GSB", "ออมสิน": "GSB"
    }
    
    for line in lines:
        if not line.strip(): 
            continue
            
        for th_name, en_abbr in bank_mapping.items():
            line = line.replace(th_name, en_abbr)
            
        # จัดการปีให้อยู่ในรูปแบบ ค.ศ. (ป้องการการลบซ้ำซ้อนหากเป็น ค.ศ. อยู่แล้ว)
        def convert_to_ce(match):
            day_month = match.group(1)
            year = int(match.group(2))
            if year > 2400: # หากเป็น พ.ศ. ให้แปลงเป็น ค.ศ.
                year -= 543
            return f"{day_month}{year}"
        
        line = re.sub(r'(\d{2}/\d{2}/)(\d{4})', convert_to_ce, line)
        parts = re.split(r'\t|\s{2,}', line.strip())
        
        if len(parts) >= 11:
            record = {
                "วันที่ทำรายการ": parts[0],
                "เวลาที่ทำรายการ": parts[1],
                "ประเภทรายการ": parts[2],
                "ช่องทาง": parts[3],
                "ชื่อธนาคารต้นทาง": parts[4],
                "หมายเลขบัญชีต้นทาง": str(parts[5]), 
                "ชื่อบัญชีต้นทาง": parts[6],
                "ชื่อธนาคารปลายทาง": parts[7],
                "หมายเลขบัญชีปลายทาง": str(parts[8]),
                "ชื่อบัญชีปลายทาง": parts[9],
                "ยอดเงิน": parts[10],
                "คงเหลือ": parts[11] if len(parts) >= 12 else "-"
            }
            records.append(record)
            
    df = pd.DataFrame(records)
    return df

# ==========================================
# ส่วนติดต่อผู้ใช้งาน (UI)
# ==========================================
st.set_page_config(page_title="ระบบวิเคราะห์รายการเดินบัญชี", layout="wide")
st.title("ระบบวิเคราะห์และสกัดข้อมูลรายการเดินบัญชี")

st.subheader("⚙️ กำหนดรูปแบบข้อความ (Template)")
default_template = "วันที่ {date} เวลา {time} มีการทำรายการ {type} จาก{src_bank} บัญชี {src_acc} ({src_name}) ไปยัง{dst_bank} บัญชี {dst_acc} ({dst_name}) ยอดเงิน {amount} บาท"
user_template = st.text_area("ตัวแปรที่รองรับ: {date}, {time}, {type}, {channel}, {src_bank}, {src_acc}, {src_name}, {dst_bank}, {dst_acc}, {dst_name}, {amount}, {balance}", value=default_template)

st.subheader("📥 นำเข้าข้อมูล")
raw_text = st.text_area("วางข้อมูลรายการเดินบัญชีลงในช่องนี้", height=200)

if st.button("ประมวลผลข้อมูล"):
    if raw_text:
        df = process_statement_data(raw_text)
        
        if not df.empty:
            st.success("ประมวลผลสำเร็จ")
            st.divider()
            
            # 1. การสร้างข้อความ Word Template
            st.subheader("📝 ข้อความสำหรับรายงานสืบสวน (Word)")
            generated_text = ""
            for _, row in df.iterrows():
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
            
            # 2. ตารางข้อมูล Excel
            st.subheader("📊 ข้อมูลตาราง (หัวตารางต้นฉบับ)")
            st.dataframe(df)
            
            col1, col2 = st.columns(2)
            
            # ส่งออกไฟล์ Excel
            with col1:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Statement')
                
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel",
                    data=excel_buffer.getvalue(),
                    file_name="Statement_Processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # ส่งออกรูปภาพ
            with col2:
                image_buffer = io.BytesIO()
                try:
                    dfi.export(df, image_buffer, table_conversion="matplotlib")
                    st.download_button(
                        label="🖼️ ดาวน์โหลดรูปภาพตาราง",
                        data=image_buffer.getvalue(),
                        file_name="Statement_Table.png",
                        mime="image/png"
                    )
                except Exception:
                    st.warning("เซิร์ฟเวอร์ไม่รองรับการส่งออกภาพตาราง โปรดใช้เครื่องมือจับภาพหน้าจอแทน")
        else:
            st.error("ไม่สามารถสกัดข้อมูลได้ โปรดตรวจสอบความถูกต้องของข้อมูลต้นฉบับ")
    else:
        st.warning("กรุณาวางข้อมูลก่อนดำเนินการ")
