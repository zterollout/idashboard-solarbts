@echo off
chcp 65001 >nul
title Update Excel Data → GitHub → Render

set PATH=%PATH%;C:\Program Files\Git\bin

echo.
echo =====================================================
echo   อัปเดต Excel Data ขึ้น Render.com
echo =====================================================
echo.

cd /d "%~dp0"

:: ── ตรวจสอบไฟล์ Excel ───────────────────────────────
if not exist "ZTE-AIS-Gulf Solar BTS 2025 Overall Progress.xlsx" (
    echo [ERROR] ไม่พบไฟล์ Excel!
    echo   กรุณาวางไฟล์ไว้ใน: %~dp0
    echo.
    pause
    exit /b 1
)
echo [OK] พบไฟล์ Excel
echo.

:: ── เช็คการเปลี่ยนแปลงก่อน add ─────────────────────
git status --porcelain "ZTE-AIS-Gulf Solar BTS 2025 Overall Progress.xlsx" "UserLogin.xlsx" > "%TEMP%\git_chk.txt"
for %%A in ("%TEMP%\git_chk.txt") do set CHKSIZE=%%~zA
if "%CHKSIZE%"=="0" (
    echo [INFO] ไม่มีการเปลี่ยนแปลง (Excel + UserLogin)
    echo        ไม่ต้อง push
    echo.
    pause
    exit /b 0
)

:: ── Git add ไฟล์ Excel + UserLogin ─────────────────
git add "ZTE-AIS-Gulf Solar BTS 2025 Overall Progress.xlsx"
git add "UserLogin.xlsx"

:: ── แสดงขนาดไฟล์ ────────────────────────────────────
for %%F in ("ZTE-AIS-Gulf Solar BTS 2025 Overall Progress.xlsx") do (
    echo [INFO] Excel size : %%~zF bytes
)
for %%F in ("UserLogin.xlsx") do (
    echo [INFO] UserLogin size: %%~zF bytes
)
echo.

:: ── Commit ───────────────────────────────────────────
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set TODAY=%%c-%%b-%%a
for /f "tokens=1-2 delims=: " %%a in ("%time%") do set NOW=%%a:%%b

git commit -m "Update Excel data [%TODAY% %NOW%]"
echo.

:: ── Push ─────────────────────────────────────────────
echo [GIT] กำลัง push ไปยัง GitHub...
git push origin main
if errorlevel 1 (
    echo.
    echo [ERROR] Push ไม่สำเร็จ
    echo   Token อาจหมดอายุ — รัน deploy_to_render.bat แล้วใส่ token ใหม่
    echo.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo   [SUCCESS] อัปเดตสำเร็จ!
echo =====================================================
echo.
echo   Render.com กำลัง reload ข้อมูลใหม่...
echo   รอประมาณ 1-2 นาที แล้วรีเฟรชเว็บ
echo.
echo   URL: https://idashboard-solarbts.onrender.com
echo.
set /p OPEN_WEB="เปิดเว็บเลย? (Y/N): "
if /i "%OPEN_WEB%"=="Y" start "" "https://idashboard-solarbts.onrender.com"
echo.
pause
