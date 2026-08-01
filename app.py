import streamlit as st
import pandas as pd
import re
import io
import dataframe_image as dfi

def process_statement_data(raw_text):
    records = []
    lines = raw_text.strip().split('\n')
    
    # แปลงชื่อธนาคารเป็นอักษรย่อภาษาอังกฤษมาตรฐาน
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
            
        # แปลงปี พ.ศ. เป็น ค.ศ.
        def convert_to_ce(match):
            day_month = match.group(1)
            year_be = int(match.group(2))
            return f"{day_month}{year_be - 543}"
        
        line = re.sub(r'(\d{2}/\d{2}/)(\d{4})', convert_to_ce, line)
        
        parts = re.split(r'\t|\s{2,}', line.strip())
        
        # รองรับกรณีมี 11 คอลัมน์ (ไม่มีคงเหลือ) หรือ 12 คอลัมน์ (มีคงเหลือ)
        if len(parts) >= 11:
            record = {
                "date": parts[0],
                "time": parts[1],
                "type": parts[2],
                "channel": parts[3],
                "src_bank": parts[4],
                "src_acc": str(parts[5]), # บังคับ String ป้องกันเลข 0 หาย
                "src_name": parts[6],
                "dst_bank": parts[7],
                "dst_acc": str(parts[8]), # บังคับ String ป้องกันเลข 0 หาย
                "dst_name": parts[9],
                "amount": parts[10],
                "balance": parts[11] if len(parts) >= 12 else "-"
            }
            records.append(record)
            
    return pd.DataFrame(records)

# ==========================================
# ส่วนติดต่อผู้ใช้งาน (User Interface)
# ==========================================
st.set_page_config(page_title="ระบบวิเคราะห์รายการเดินบัญชี", layout="wide")
st.title("ระบบวิเคราะห์และสกัดข้อมูลรายการเดินบัญชี")

# ส่วนที่ 1: ตั้งค่า Template
st.subheader("⚙️ กำหนดรูปแบบข้อความ (Template)")
default_template = "วันที่ {date} เวลา {time} มีการทำรายการ {type} จากธนาคาร {src_bank} บัญชี {src_acc} ({src_name}) ไปยังธนาคาร {dst_bank} บัญชี {dst_acc} ({dst_name}) ยอดเงิน {amount} บาท"
user_template = st.text_area("ใช้ตัวแปรในวงเล็บปีกกาเพื่อจัดรูปประโยค", value=default_template)

st.markdown("""
*ตัวแปรที่ใช้ได้:* `{date}`, `{time}`, `{type}`, `{channel}`, `{src_bank}`, `{src_acc}`, `{src_name}`, `{dst_bank}`, `{dst_acc}`, `{dst_name}`, `{amount}`, `{balance}`
""")

# ส่วนที่ 2: รับข้อมูล
st.subheader("📥 นำเข้าข้อมูล")
raw_text = st.text_area("วางข้อมูลรายการเดินบัญชีที่คัดลอกมา (Copy & Paste) ลงในช่องนี้", height=200)

if st.button("ประมวลผลข้อมูล"):
    if raw_text:
        df = process_statement_data(raw_text)
        
        if not df.empty:
            st.success("ประมวลผลสำเร็จ")
            st.divider()
            
            # ส่วนแสดงผลข้อความตาม Template
            st.subheader("📝 ข้อความสำหรับนำไปใช้วางใน Word")
            generated_text = ""
            for _, row in df.iterrows():
                try:
                    # นำค่าจากตารางมาแทนที่ใน Template
                    row_text = user_template.format(**row.to_dict())
                    generated_text += row_text + "\n"
                except KeyError as e:
                    st.error(f"เกิดข้อผิดพลาด: ไม่พบตัวแปร {e} ในระบบ กรุณาตรวจสอบ Template")
                    break
            
            if generated_text:
                st.text_area("คัดลอกข้อความด้านล่าง:", value=generated_text, height=200)
            
            st.divider()
            
            # ส่วนแสดงผลตารางและส่งออก
            st.subheader("📊 ตารางข้อมูลและการส่งออก")
            st.dataframe(df)
            
            col1, col2 = st.columns(2)
            
            # ส่งออกเป็น Excel
            with col1:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Statement')
                
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
                    data=excel_buffer.getvalue(),
                    file_name="Statement_Processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # ส่งออกเป็นรูปภาพ
            with col2:
                # สร้างรูปภาพตารางด้วย dataframe_image
                image_buffer = io.BytesIO()
                try:
                    dfi.export(df, image_buffer, table_conversion="matplotlib")
                    st.download_button(
                        label="🖼️ ดาวน์โหลดรูปภาพตาราง (.png)",
                        data=image_buffer.getvalue(),
                        file_name="Statement_Table.png",
                        mime="image/png"
                    )
                except Exception as e:
                    st.warning("ระบบเซิร์ฟเวอร์อาจไม่รองรับการแปลงภาพตารางโดยตรง แนะนำให้ใช้ Snipping Tool ในการจับภาพหน้าจอแทน")
                    
        else:
            st.error("ไม่สามารถสกัดข้อมูลได้ กรุณาตรวจสอบว่ารูปแบบข้อมูลตรงกับที่กำหนดหรือไม่")
    else:
        st.warning("กรุณาวางข้อมูลก่อนกดประมวลผล")
