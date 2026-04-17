@echo off
chcp 65001 >nul
title Setup Git — First Time Only

echo.
echo =====================================================
echo   Git First-Time Setup (ทำครั้งเดียวเท่านั้น)
echo =====================================================
echo.

where git >nul 2>&1
if errorlevel 1 (
    echo [ERROR] ยังไม่มี Git — กรุณาติดตั้งก่อน
    echo   https://git-scm.com/download/win
    pause
    exit /b 1
)

echo [1/4] ตั้งชื่อ Git user
set /p GIT_NAME="ชื่อ (เช่น John Doe): "
git config --global user.name "%GIT_NAME%"

echo.
echo [2/4] ตั้ง Email Git user
set /p GIT_EMAIL="Email (ใช้ email เดียวกับ GitHub): "
git config --global user.email "%GIT_EMAIL%"

echo.
echo [3/4] Init repository
if not exist ".git" (
    git init
    git branch -M main
    echo [OK] สร้าง repo ใหม่แล้ว
) else (
    echo [OK] มี repo อยู่แล้ว
)

echo.
echo [4/4] เชื่อม GitHub remote
echo   - ไปที่ https://github.com/new
echo   - สร้าง repository ใหม่ ชื่อ idashboard-solarbts
echo   - เลือก Private (แนะนำ — ข้อมูลมีความลับ)
echo   - ไม่ต้อง init with README
echo.
set /p REPO_URL="วาง GitHub repo URL (เช่น https://github.com/username/idashboard-solarbts.git): "
git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%

echo.
echo =====================================================
echo   Setup เสร็จแล้ว!
echo   รัน deploy_to_render.bat เพื่อ deploy ครั้งแรก
echo =====================================================
echo.
pause
