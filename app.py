import streamlit as st
import pandas as pd
import re
import io
import dataframe_image as dfi

def process_statement_data(raw_text):
    """
    ฟังก์ชันสกัดและทำความสะอาดข้อมูลรายการเดินบัญชี
    """
    records = []
    lines = raw_text.strip().split('\n')
    
    # 1. กำหนดพจนานุกรมสำหรับแปลงชื่อธนาคารเป็นชื่อย่อภาษาอังกฤษมาตรฐาน
    bank_mapping = {
        "ธนาคารกสิกรไทย": "KBANK", "กสิกรไทย": "KBANK",
        "ธนาคารไทยพาณิชย์": "SCB", "ไทยพาณิชย์": "SCB",
        "ธนาคารกรุงเทพ": "BBL", "กรุงเทพ": "BBL",
        "ธนาคารกรุงไทย": "KTB", "กรุงไทย": "KTB"
    }
    
    for line in lines:
        if not line.strip(): 
            continue
            
        # แปลงชื่อธนาคาร
        for th_name, en_abbr in bank_mapping.items():
            line = line.replace(th_name, en_abbr)
            
        # 2. แปลงรูปแบบวันที่จาก พ.ศ. เป็น ค.ศ. (dd/mm/yyyy)
        def convert_to_ce(match):
            day_month = match.group(1)
            year_be = int(match.group(2))
            return f"{day_month}{year_be - 543}"
        
        line = re.sub(r'(\d{2}/\d{2}/)(\d{4})', convert_to_ce, line)
        
        # แยกข้อมูลด้วยช่องว่างหรือ Tab (อ้างอิงจากการคัดลอกตาราง)
        parts = re.split(r'\t|\s{2,}', line.strip())
        
        if len(parts) >= 12:
            record = {
                "วันที่ทำรายการ": parts[0],
                "เวลาที่ทำรายการ": parts[1],
                "ประเภทรายการ": parts[2],
                "ช่องทาง": parts[3],
                "ชื่อธนาคารต้นทาง": parts[4],
                "หมายเลขบัญชีต้นทาง": str(parts[5]), # 3. บังคับเป็น String ป้องกันเลข 0 หาย
                "ชื่อบัญชีต้นทาง": parts[6],
                "ชื่อธนาคารปลายทาง": parts[7],
                "หมายเลขบัญชีปลายทาง": str(parts[8]), # 3. บังคับเป็น String ป้องกันเลข 0 หาย
                "ชื่อบัญชีปลายทาง": parts[9],
                "ยอดเงิน": parts[10],
                "คงเหลือ": parts[11]
            }
            records.append(record)
            
    return pd.DataFrame(records)

def generate_word_template(df):
    """
    ฟังก์ชันสร้างข้อความ Template สำหรับนำไปวางใน Word
    """
    text_output = ""
    for _, row in df.iterrows():
        text_output += f"วันที่: {row['วันที่ทำรายการ']} เวลา: {row['เวลาที่ทำรายการ']}\n"
        text_output += f"รายการ: {row['ประเภทรายการ']} ผ่าน {row['ช่องทาง']}\n"
        text_output += f"ต้นทาง: {row['ชื่อธนาคารต้นทาง']} {row['หมายเลขบัญชีต้นทาง']} ({row['ชื่อบัญชีต้นทาง']})\n"
        text_output += f"ปลายทาง: {row['ชื่อธนาคารปลายทาง']} {row['หมายเลขบัญชีปลายทาง']} ({row['ชื่อบัญชีปลายทาง']})\n"
        text_output += f"ยอดเงิน: {row['ยอดเงิน']} | คงเหลือ: {row['คงเหลือ']}\n"
        text_output += "=" * 60 + "\n"
    return text_output

# ==========================================
# ส่วนติดต่อผู้ใช้งาน (User Interface)
# ==========================================
st.set_page_config(page_title="ระบบวิเคราะห์รายการเดินบัญชี", layout="wide")
st.title("ระบบวิเคราะห์และสกัดข้อมูลรายการเดินบัญชี")

raw_text = st.text_area("วางข้อมูลรายการเดินบัญชี (คัดลอกตารางจาก PDF/Excel มาวางในช่องนี้)", height=250)

if st.button("ดำเนินการประมวลผล"):
    if raw_text:
        df = process_statement_data(raw_text)
        
        if not df.empty:
            st.success("ประมวลผลข้อมูลเสร็จสิ้น")
            
            col1, col2 = st.columns(2)
            
            # ส่วนที่ 1: ข้อความสำหรับ Word
            with col1:
                st.subheader("1. ข้อความ Template สำหรับ Word")
                text_template = generate_word_template(df)
                st.text_area("สามารถคัดลอกข้อความด้านล่างได้ทันที", value=text_template, height=350)
            
            # ส่วนที่ 2: ตารางสำหรับ Excel และรูปภาพ
            with col2:
                st.subheader("2. ข้อมูลตาราง")
                st.dataframe(df)
                
                # เตรียมไฟล์ Excel
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='StatementData')
                excel_data = excel_buffer.getvalue()
                
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel",
                    data=excel_data,
                    file_name="Statement_Processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # เตรียมไฟล์ภาพตาราง
                st.info("💡 หากต้องการภาพตาราง สามารถใช้เครื่องมือ Snipping Tool แคปเจอร์จากหน้าจอนี้ หรือใช้ไลบรารี dfi.export(df, 'table.png') เพื่อบันทึกเป็นภาพอัตโนมัติ")
                
        else:
            st.error("ไม่สามารถสกัดข้อมูลได้ กรุณาตรวจสอบว่ารูปแบบข้อมูลที่วางตรงกับโครงสร้างที่กำหนดหรือไม่")
    else:
        st.warning("กรุณาวางข้อมูลก่อนกดดำเนินการ")