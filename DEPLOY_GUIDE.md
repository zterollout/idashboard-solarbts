# 📋 คู่มือ Deploy บน Render.com
## โปรเจกต์: idashboard-solarbts

---

## 🔧 สิ่งที่ต้องเตรียม (ทำครั้งเดียว)

### 1. ติดตั้ง Git
- ดาวน์โหลด: https://git-scm.com/download/win
- ติดตั้งแบบ default ทั้งหมด
- เปิด PowerShell ใหม่แล้วทดสอบ: `git --version`

### 2. สร้างบัญชี GitHub
- ไปที่: https://github.com
- Sign up → ใช้ email ของคุณ
- แนะนำเปิดใช้ 2FA เพื่อความปลอดภัย

### 3. สร้าง GitHub Repository
1. ไปที่ https://github.com/new
2. ตั้งชื่อ: **`idashboard-solarbts`**
3. เลือก **Private** ⚠️ (แนะนำ เพราะมีข้อมูลโครงการ)
4. **ไม่ต้อง** check "Add README", ".gitignore", "license"
5. กด **Create repository**
6. Copy URL รูปแบบ: `https://github.com/YOUR_USERNAME/idashboard-solarbts.git`

### 4. สร้างบัญชี Render.com
- ไปที่: https://render.com
- Sign up → **Continue with GitHub** (แนะนำ)

---

## 🚀 ขั้นตอน Deploy ครั้งแรก

### ขั้นตอนที่ 1 — Setup Git (ทำครั้งเดียว)
```
รัน: git_first_setup.bat
```
- ใส่ชื่อ, email, และ GitHub URL

### ขั้นตอนที่ 2 — Push โค้ดขึ้น GitHub
```
รัน: deploy_to_render.bat
```
- ใส่ commit message เช่น "Initial deploy"
- ระบบจะ push ทุกไฟล์รวมถึง Excel ขึ้น GitHub

### ขั้นตอนที่ 3 — สร้าง Web Service บน Render
1. ไปที่ https://dashboard.render.com
2. กด **New +** → **Web Service**
3. เลือก **Connect a repository** → เลือก `idashboard-solarbts`
4. ตั้งค่าดังนี้:

| Field | ค่า |
|-------|-----|
| Name | `idashboard-solarbts` |
| Region | Singapore (ใกล้ไทยสุด) |
| Branch | `main` |
| Runtime | `Python 3` |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn main:app --host 0.0.0.0 --port $PORT` |
| Plan | **Free** |

5. กด **Create Web Service**
6. รอ build ประมาณ **3-5 นาที**
7. เมื่อ status เป็น **Live** → เข้าเว็บได้ที่:
   **https://idashboard-solarbts.onrender.com**

---

## 🔄 อัปเดตข้อมูล Excel

ทุกครั้งที่อัปเดตข้อมูล:
```
รัน: deploy_to_render.bat
```
- ใส่ message เช่น "Update Excel data 17-Apr-2026"
- Render จะ **auto-deploy** ภายใน 2-3 นาที

---

## ⚠️ ข้อจำกัด Free Tier

| ข้อจำกัด | รายละเอียด |
|----------|------------|
| Sleep mode | หยุดทำงานหลัง **15 นาที** ไม่มีคนใช้ |
| Wake up time | request แรกใช้เวลา **~30 วินาที** |
| RAM | **512 MB** |
| Build minutes | **500 นาที/เดือน** |
| Bandwidth | **100 GB/เดือน** |

> 💡 **เคล็ดลับ**: ถ้าไม่อยากให้ sleep ให้ใช้ UptimeRobot (ฟรี) ping ทุก 14 นาที

---

## 🔐 ความปลอดภัย

- Repository ตั้งเป็น **Private** → ไม่มีใครเห็นโค้ดและข้อมูล
- เว็บมีระบบ **Login** อยู่แล้ว → ต้องใส่ user/password ก่อนเข้า
- ผู้ใช้งานจัดการได้ใน `UserLogin.xlsx`

---

## 🆘 แก้ปัญหาที่พบบ่อย

### Push ไม่สำเร็จ (Authentication failed)
GitHub ไม่รับ password ธรรมดาแล้ว ต้องใช้ **Personal Access Token**:
1. ไปที่ GitHub → Settings → Developer Settings → Personal access tokens
2. Generate new token (classic) → เลือก scope: `repo`
3. Copy token → ใช้แทน password ตอน push

### Build ล้มเหลวบน Render
- ตรวจดู Build log ใน Render dashboard
- มักเกิดจาก Python version ไม่ตรง → เพิ่มไฟล์ `runtime.txt`:
  ```
  python-3.11.9
  ```

### เว็บโหลดช้ามาก
- เกิดจาก Free tier sleep → ปกติ
- แก้ได้โดย upgrade เป็น Starter plan ($7/เดือน) หรือใช้ UptimeRobot

### Excel ไม่โหลด
- ตรวจสอบชื่อไฟล์ว่าตรงกับที่ตั้งใน render.yaml
- ตรวจ `EXCEL_FILE` environment variable ใน Render dashboard

---

## 📁 โครงสร้างไฟล์ที่ push ขึ้น GitHub

```
idashboard-solarbts/
├── main.py                                          ← Backend API
├── requirements.txt                                 ← Python dependencies
├── render.yaml                                      ← Render config
├── .gitignore                                       ← ไฟล์ที่ไม่ push
├── UserLogin.xlsx                                   ← รายชื่อผู้ใช้
├── ZTE-AIS-Gulf Solar BTS 2025 Overall Progress.xlsx ← ข้อมูลหลัก
├── HLP.csv                                          ← ข้อมูล HLP
└── static/
    ├── index.html                                   ← หน้า Dashboard
    ├── login.html                                   ← หน้า Login
    └── admin.html                                   ← หน้า Admin
```

---

*อัปเดตล่าสุด: เมษายน 2026*
