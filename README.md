# 📘 คู่มือการใช้งาน ZTE-AIS Gulf Solar BTS Dashboard 2025

---

## ⚠️ ข้อมูลสำคัญ — Data Source (อย่าลืม!)

| รายละเอียด | ค่า |
|---|---|
| **ไฟล์ Excel** | `ZTE-AIS-Gulf Solar BTS 2025 Overall Progress.xlsx` |
| **Sheet name** | `DATA` |
| **Header row** | แถวที่ **6** ใน Excel → `header=5` (0-indexed ใน pandas) |
| **โค้ดใน main.py** | `HEADER_ROW = 5` และ `sheet_name="DATA"` |

### คอลัมน์หลักที่ใช้ใน Dashboard

| คอลัมน์ใน Excel | ใช้ใน Dashboard |
|---|---|
| `RFI Status` | Go / No Go / Wait Review |
| `ETSS Status` | ETSS Approved (value = "Approved") |
| `AIS Confirm Solution Date` | นับ rows ที่มีวันที่ (not null) |
| `PBOM Confirm Actual` | นับ rows ที่มีค่า (not null/blank) |
| `iEPMS PBOM Status` | PBOM pie chart (Confirm / To be check ฯลฯ) |
| `BOM by site Status` | Released / Draft / Changing = Actual |
| `Issue Grouping` | No Go Issue Grouping table |
| `Region` | แบ่งตาม region ทุกตาราง |

---

## 📁 โครงสร้างไฟล์

```
d:\PythonWork\
├── main.py                                      ← FastAPI backend (server หลัก)
├── requirements.txt                             ← Python dependencies
├── render.yaml                                  ← Render.com deploy config
├── .gitignore                                   ← ไฟล์ที่ไม่ต้องการ push
├── start_dashboard.bat                          ← ไฟล์รันด่วน (ดับเบิ้ลคลิกได้เลย)
├── static\
│   ├── index.html                              ← หน้า Dashboard (Frontend)
│   └── login.html                              ← หน้า Login
├── UserLogin.xlsx                               ← ข้อมูล User สำหรับ Login
└── ZTE-AIS-Gulf Solar BTS 2025 Overall Progress.xlsx  ← ไฟล์ข้อมูล Excel (sheet "DATA" and header row start row 6)
```

---

## ☁️ Deploy บน Cloud (Render.com) — ฟรี ไม่จำกัดเวลา

### ขั้นตอนที่ 1 — สร้าง GitHub Repository

1. ไปที่ [github.com](https://github.com) → สร้าง account (ถ้ายังไม่มี)
2. คลิก **New Repository** → ตั้งชื่อ เช่น `solar-bts-dashboard`
3. เลือก **Private** (เพื่อความปลอดภัยของข้อมูล)
4. คลิก **Create Repository**

ที่เครื่อง Windows — รันคำสั่งต่อไปนี้ใน PowerShell:

```powershell
cd d:\PythonWork

# ครั้งแรก: ติดตั้ง git ถ้ายังไม่มี → https://git-scm.com/download/win

git init
git add .
git commit -m "Initial commit: Solar BTS Dashboard"
git remote add origin https://github.com/<YOUR_USERNAME>/solar-bts-dashboard.git
git branch -M main
git push -u origin main
```

---

### ขั้นตอนที่ 2 — Deploy บน Render.com

1. ไปที่ [render.com](https://render.com) → สมัคร/Login ด้วย GitHub account
2. คลิก **New → Web Service**
3. เลือก Repository `solar-bts-dashboard` → คลิก **Connect**
4. กรอกข้อมูล:
   - **Name**: `solar-bts-dashboard`
   - **Region**: Singapore (ใกล้ไทยสุด)
   - **Branch**: `main`
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`
   - **Plan**: **Free**
5. คลิก **Create Web Service**
6. รอ ~3-5 นาที → ได้ URL เช่น `https://solar-bts-dashboard.onrender.com`

---

### ขั้นตอนที่ 3 — อัปเดตข้อมูล Excel บน Cloud

เมื่อต้องการอัปเดตข้อมูล Excel:

```powershell
cd d:\PythonWork
# แทนที่ไฟล์ Excel ใหม่แล้ว:
git add "ZTE-AIS-Gulf Solar BTS 2025 Overall Progress.xlsx"
git add UserLogin.xlsx
git commit -m "Update data: $(Get-Date -Format 'dd-MMM-yyyy')"
git push
```
Render จะ **auto-redeploy** ภายใน ~2 นาที

---

### ⚠️ ข้อจำกัด Free Tier ของ Render

| หัวข้อ | รายละเอียด |
|---|---|
| **Sleep** | Service จะ sleep หลังไม่มี traffic 15 นาที |
| **Wake up** | ครั้งแรกหลัง sleep ใช้เวลา ~30-50 วินาที |
| **RAM** | 512 MB (เพียงพอสำหรับ Excel ~50MB) |
| **Bandwidth** | 100 GB/เดือน (มากเกินพอ) |
| **ราคา** | **ฟรี** ตลอดไป ตราบที่ยังใช้ Free plan |

> **Tip**: ถ้าไม่อยากให้ sleep ให้ใช้ [UptimeRobot](https://uptimerobot.com) (ฟรี) ping ทุก 5 นาที

---

## 🚀 วิธีรัน Dashboard บนเครื่อง Local

### ✅ แบบที่ 1 — ดับเบิ้ลคลิก (ง่ายที่สุด)

1. ไปที่ `d:\PythonWork\`
2. ดับเบิ้ลคลิกที่ไฟล์ **`start_dashboard.bat`**
3. Browser จะเปิดที่ `http://127.0.0.1:8000`

---

### ✅ แบบที่ 2 — รันผ่าน Terminal

```powershell
cd d:\PythonWork
.venv\Scripts\uvicorn.exe main:app --reload --port 8000
```

---

## 🔄 การอัปเดตข้อมูล Excel (Local)

1. **ปิดไฟล์ Excel** ก่อน
2. แทนที่ไฟล์ `ZTE-AIS-Gulf Solar BTS 2025 Overall Progress.xlsx`
3. **Refresh browser** (`F5`) — ระบบโหลดข้อมูลใหม่อัตโนมัติ

---

## 👤 ระบบ Login

| Username | Password | Role |
|---|---|---|
| `admin` | `admin` | Admin |
| `user1` | `user1` | Member |

เพิ่ม/แก้ไข user ได้ที่ไฟล์ `UserLogin.xlsx` (Sheet1 columns: User, Password, Role)

---

## 📦 ติดตั้ง Dependencies ใหม่

```powershell
cd d:\PythonWork
python -m venv .venv
.venv\Scripts\pip.exe install -r requirements.txt
```

---

*Last updated: April 2026 | ZTE-AIS Gulf Solar BTS 2025 Dashboard v2.0*
