import streamlit as st
import pandas as pd
import json
import os
import datetime
from io import BytesIO

from src.parser import load_and_validate_inputs, parse_master_data_file
from src.transformer import transform_to_autoline_data
from src.exporter import generate_output_excel

st.set_page_config(
    page_title="Autoline Excel Interface Bridge",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E88E5;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #546E7A;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Load default config & templates
base_dir = os.path.dirname(__file__)
config_path = os.path.join(base_dir, "config", "default_settings.json")
template_path = os.path.join(base_dir, "data", "output_template.xlsx")
master_template_path = os.path.join(base_dir, "data", "master_data_template.xlsx")

def load_settings():
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

if "settings" not in st.session_state:
    st.session_state.settings = load_settings()

# Sidebar: GL Configuration & Master Settings
with st.sidebar:
    st.header("⚙️ การตั้งค่าผังบัญชี (GL Mapping)")
    st.caption("ปรับเปลี่ยนรหัสบัญชีและค่าคงที่สำหรับส่งออก Autoline")
    
    with st.expander("📌 ข้อมูลทั่วไป (General)", expanded=True):
        st.session_state.settings["journal_desc"] = st.text_input(
            "คำอธิบายสมุดรายวัน (JNLDESC)",
            value=st.session_state.settings.get("journal_desc", "After Sale")
        )
        st.session_state.settings["src_branch"] = st.text_input(
            "รหัสสาขาเริ่มต้น (SRCBRANCH)",
            value=st.session_state.settings.get("src_branch", "0001")
        )
        st.session_state.settings["subaccount_default"] = st.text_input(
            "รหัสลูกหนี้เริ่มต้น (SUBACCOUNT)",
            value=st.session_state.settings.get("subaccount_default", "A0011")
        )
        
    with st.expander("🏦 ลูกหนี้การค้า (AR Debit Line)", expanded=False):
        st.session_state.settings["ar_gl_code"] = st.text_input(
            "ผังบัญชีลูกหนี้ (GLCODE)",
            value=str(st.session_state.settings.get("ar_gl_code", "11311001"))
        )
        st.session_state.settings["ar_department"] = st.text_input(
            "แผนกลูกหนี้ (DEPARTMENT)",
            value=str(st.session_state.settings.get("ar_department", "0000"))
        )
        
    with st.expander("🔩 หมวดอะไหล่ (Parts - P)", expanded=False):
        p_cfg = st.session_state.settings.get("parts", {})
        p_cfg["gl_code"] = st.text_input("GL Code อะไหล่", value=str(p_cfg.get("gl_code", "41211004")))
        p_cfg["department"] = st.text_input("แผนกอะไหล่", value=str(p_cfg.get("department", "4002")))
        p_cfg["partfran"] = st.text_input("PARTFRAN", value=str(p_cfg.get("partfran", "J")))
        p_cfg["partprod"] = st.text_input("PARTPROD", value=str(p_cfg.get("partprod", "A")))
        st.session_state.settings["parts"] = p_cfg
        
    with st.expander("🔧 หมวดค่าแรง (Labor - L)", expanded=False):
        l_cfg = st.session_state.settings.get("labor", {})
        l_cfg["gl_code"] = st.text_input("GL Code ค่าแรง", value=str(l_cfg.get("gl_code", "41111001")))
        l_cfg["department"] = st.text_input("แผนกค่าแรง", value=str(l_cfg.get("department", "4001")))
        l_cfg["servprod"] = st.text_input("SERVPROD", value=str(l_cfg.get("servprod", "A")))
        l_cfg["servicefranc"] = st.text_input("SERVICEFRANC", value=str(l_cfg.get("servicefranc", "JEEP")))
        st.session_state.settings["labor"] = l_cfg
        
    with st.expander("🛠️ หมวดบริการภายนอก (Sublet - S)", expanded=False):
        s_cfg = st.session_state.settings.get("sublet", {})
        s_cfg["gl_code"] = st.text_input("GL Code Sublet", value=str(s_cfg.get("gl_code", "41311001")))
        s_cfg["department"] = st.text_input("แผนก Sublet", value=str(s_cfg.get("department", "4003")))
        s_cfg["servprod"] = st.text_input("SERVPROD (Sublet)", value=str(s_cfg.get("servprod", "A")))
        s_cfg["servicefranc"] = st.text_input("SERVICEFRANC (Sublet)", value=str(s_cfg.get("servicefranc", "JEEP")))
        st.session_state.settings["sublet"] = s_cfg

    if st.button("💾 บันทึกการตั้งค่า (Save Settings)", use_container_width=True):
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(st.session_state.settings, f, ensure_ascii=False, indent=2)
        st.success("บันทึกการตั้งค่าเรียบร้อยแล้ว!")

# Main App Header
st.markdown('<div class="main-header">🚗 Autoline Excel Interface Bridge</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">ระบบแปลงข้อมูลบิลบริการและอะไหล่เข้าสู่รูปแบบ Autoline (CDK) AR/AP Journal Template</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🚀 แปลงข้อมูล & ดาวน์โหลด (Transform & Export)", "📑 จัดการ Master Data (Master Mapping)"])

with tab1:
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.subheader("1. อัปโหลดไฟล์ HEADER")
        st.caption("ไฟล์ที่มีเลขที่บิล, วันที่, รหัสลูกค้า, และยอดรวม (เช่น Input Template (HEADER).xlsx)")
        header_file = st.file_uploader("เลือกไฟล์ HEADER (.xlsx)", type=["xlsx", "xls"], key="header_upload")
        
    with col_up2:
        st.subheader("2. อัปโหลดไฟล์ DETAIL")
        st.caption("ไฟล์ที่มีรายการแยกตามหมวด P, L, S (เช่น Input Template (DETAIL).xlsx)")
        detail_file = st.file_uploader("เลือกไฟล์ DETAIL (.xlsx)", type=["xlsx", "xls"], key="detail_upload")
        
    # Optional Master Data upload
    master_file = st.file_uploader(
        "📎 (ทางเลือก) อัปโหลดไฟล์ Master Data เพื่อ Map รหัสลูกหนี้/สาขาอัตโนมัติ (.xlsx)",
        type=["xlsx", "xls"],
        key="master_upload"
    )
    master_dict = None
    if master_file:
        try:
            master_dict = parse_master_data_file(master_file)
            st.success(f"✅ โหลด Master Data สำเร็จ! (พบข้อมูลลูกค้า {len(master_dict.get('customers', {}))} ราย, สาขา {len(master_dict.get('branches', {}))} สาขา)")
        except Exception as e:
            st.warning(f"ไม่สามารถอ่านไฟล์ Master Data ได้: {str(e)}")

    if header_file and detail_file:
        try:
            with st.spinner("กำลังอ่านและตรวจสอบความถูกต้องของข้อมูล..."):
                df_header, df_detail = load_and_validate_inputs(header_file, detail_file)
                rows_to_write, summary_stats, preview_df = transform_to_autoline_data(
                    df_header, df_detail, st.session_state.settings, master_dict=master_dict
                )
                
            st.success("✅ ประมวลผลและจับคู่ข้อมูลสำเร็จ!")
            
            # KPI Dashboard
            st.markdown("### 📊 สรุปภาพรวมข้อมูลและตรวจความสมดุลทางบัญชี (Balance Check)")
            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            kpi1.metric("จำนวนใบแจ้งหนี้", f"{summary_stats['total_invoices']:,} ใบ")
            kpi2.metric("ยอดขายสุทธิ (Credit)", f"{summary_stats['total_credit']:,.2f} ฿")
            kpi3.metric("ภาษี 7% (VAT)", f"{summary_stats['total_tax']:,.2f} ฿")
            kpi4.metric("ยอดลูกหนี้รวม (Debit)", f"{summary_stats['total_debit']:,.2f} ฿")
            
            if summary_stats["is_balanced"]:
                kpi5.metric("สถานะดุลบัญชี", "สมดุล 100% ✅", delta="Dr = Cr + Tax")
            else:
                kpi5.metric("สถานะดุลบัญชี", "ยอดไม่ดุล ⚠️", delta=f"Diff: {summary_stats['difference']}")

            # Export Button
            st.markdown("---")
            col_dl1, col_dl2 = st.columns([2, 1])
            with col_dl1:
                st.subheader("📥 ดาวน์โหลดไฟล์สำหรับ Autoline")
                st.caption("ไฟล์ Excel ในโครงสร้าง Autoline AR/AP พร้อมนำเข้าทันที (DEBIT แถว 9, CREDIT แถว 10, เว้นแถว 8)")
            with col_dl2:
                excel_output = generate_output_excel(template_path, rows_to_write)
                now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="⬇️ ดาวน์โหลด Autoline_Import.xlsx",
                    data=excel_output,
                    file_name=f"Autoline_Import_{now_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            # Table Preview
            st.markdown("### 🔍 ตัวอย่างข้อมูลที่แปลงแล้ว (Data Preview)")
            search_inv = st.text_input("ค้นหาเลขที่บิล (Filter by Invoice Number)", placeholder="พิมพ์เลขที่บิล เช่น 01S26080001")
            
            display_df = preview_df
            if search_inv.strip():
                display_df = preview_df[preview_df["Invoice"].astype(str).str.contains(search_inv.strip(), case=False, na=False)]
                
            st.dataframe(display_df, use_container_width=True, height=450)
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ HEADER และ DETAIL ด้านบน หรือทดสอบด้วยข้อมูลตัวอย่างในเครื่อง")
        
        sample_h_path = r"C:/Users/Nawarutte.Non/.gemini/antigravity/scratch/Autoline Interface/Input Template (HEADER).xlsx"
        sample_d_path = r"C:/Users/Nawarutte.Non/.gemini/antigravity/scratch/Autoline Interface/Input Template (DETAIL).xlsx"
        
        if os.path.exists(sample_h_path) and os.path.exists(sample_d_path):
            if st.button("🧪 ทดสอบด้วยไฟล์ตัวอย่างในเครื่องทันที (Load Sample)"):
                df_header, df_detail = load_and_validate_inputs(sample_h_path, sample_d_path)
                rows_to_write, summary_stats, preview_df = transform_to_autoline_data(
                    df_header, df_detail, st.session_state.settings
                )
                st.success("✅ โหลดตัวอย่างและแปลงข้อมูลเรียบร้อย!")
                st.dataframe(preview_df.head(20), use_container_width=True)

with tab2:
    st.header("📑 เทมเพลตไฟล์ Master Data")
    st.markdown("""
    คุณสามารถเตรียมไฟล์ Excel Master Data สำหรับจับคู่รหัสลูกหนี้ สาขา และผังบัญชีได้ โดยไฟล์มี 3 แผ่นงาน (Sheets):
    1. **Customer_Subaccount:** จับคู่ `customer_code` หรือ `tax_ID` กับ `autoline_subaccount`
    2. **Branch_Mapping:** จับคู่ `dealer_prefix` หรือ `branch_name` กับ `autoline_branch_code`
    3. **GL_Mapping:** จับคู่ `category` (P, L, S, AR) กับ `gl_code`, `department`, `partfran`, `servprod`
    """)
    
    if os.path.exists(master_template_path):
        with open(master_template_path, "rb") as f:
            st.download_button(
                label="⬇️ ดาวน์โหลดเทมเพลต Master Data (Excel)",
                data=f.read(),
                file_name="Master_Data_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
