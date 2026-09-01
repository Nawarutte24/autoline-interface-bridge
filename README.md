# 🚗 Autoline Excel Interface Bridge (Streamlit Web App)

เว็บแอปพลิเคชันสำหรับแปลงข้อมูลไฟล์ Excel ส่งออกจากระบบศูนย์บริการ (DMS) ไปเป็นรูปแบบ **Autoline (CDK) AR/AP Journal Import Template** โดยอัตโนมัติ

---

## 🌟 ฟีเจอร์หลัก (Features)

- **Auto-Join Header & Detail:** เชื่อมโยงไฟล์ `HEADER` (ข้อมูลวันที่ ลูกค้า ยอดรวม) และ `DETAIL` (รายการแยกหมวด อะไหล่ P, ค่าแรง L, บริการ S) ด้วย `invoice_number` 100%
- **Exact Autoline Layout:** จัดรูปแบบตรงตาม `Output Template (NEW).xlsx`
  - แถว 1-6 คงโครงสร้างและหัวตารางเดิมครบถ้วน
  - แถว 7 เป็น Document Header
  - แถว 8 เป็น Blank Row เว้นบรรทัดตามแบบฟอร์ม
  - แถว 9 เป็น Line 1: **DEBIT ลูกหนี้การค้า (รวม VAT)**
  - แถว 10+ เป็น Line 2..N: **CREDIT รายได้แยกตามหมวด (อะไหล่, ค่าแรง, Sublet)**
- **Accounting Balance Integrity:** ตรวจสอบยอดเดบิตและเครดิตอัตโนมัติ ($$Debit = Credit + Tax$$)
- **Interactive GL & Master Config:** แถบด้านข้างสำหรับปรับเปลี่ยน GL Code, แผนก, Subaccount, Franchise, Product Codes พร้อมปุ่ม Save Configuration
- **Live Search & Filter Preview:** ค้นหาและดูตัวอย่างตารางผลลัพธ์ก่อนดาวน์โหลด
- **One-Click Excel Export:** ดาวน์โหลดไฟล์ `.xlsx` ที่พร้อมนำเข้า Autoline ได้ทันที

---

## 🚀 การติดตั้งและรันในเครื่อง (Local Run)

```bash
# 1. ไปที่โฟลเดอร์โปรเจกต์
cd "C:\Users\Nawarutte.Non\.gemini\antigravity\scratch\autoline-interface-app"

# 2. ติดตั้ง Dependencies
pip install -r requirements.txt

# 3. รัน Web App
streamlit run app.py
```

---

## ☁️ การนำขึ้น GitHub และ Deploy บน Streamlit Community Cloud

### 1. Push โค้ดขึ้น GitHub
```bash
git init
git add .
git commit -m "Initial commit: Autoline Interface Bridge App"
git branch -M main
git remote add origin https://github.com/<YOUR_GITHUB_USERNAME>/autoline-interface-bridge.git
git push -u origin main
```

### 2. Deploy บน Streamlit Cloud (ฟรี)
1. เข้าไปที่ [share.streamlit.io](https://share.streamlit.io) และ Login ด้วย GitHub
2. กด **Create app**
3. เลือก Repository: `autoline-interface-bridge`
4. Main file path: `app.py`
5. กด **Deploy!**
