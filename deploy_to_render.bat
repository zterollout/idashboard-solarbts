@echo off
chcp 65001 >nul
title Deploy to Render.com — idashboard-solarbts

set REPO_URL=https://github.com/zterollout/idashboard-solarbts.git
set SITE_URL=https://idashboard-solarbts.onrender.com

echo.
echo =====================================================
echo   Deploy to Render.com : idashboard-solarbts
echo   Repo : %REPO_URL%
echo =====================================================
echo.

:: ── ตรวจสอบ Git ──────────────────────────────────────
where git >nul 2>&1
if errorlevel 1 (
    echo [ERROR] ไม่พบ Git บนเครื่องนี้
    echo.
    echo กรุณาติดตั้ง Git ก่อน:
    echo   https://git-scm.com/download/win
    echo.
    pause
    exit /b 1
)
echo [OK] พบ Git
git --version
echo.

:: ── Init repo ถ้ายังไม่มี ─────────────────────────────
if not exist ".git" (
    echo [INIT] สร้าง Git repository ใหม่...
    git init
    git branch -M main
    git remote add origin %REPO_URL%
    echo [OK] เชื่อม remote origin แล้ว
    echo.
) else (
    :: ตรวจสอบและอัปเดต remote URL ให้ถูกต้องเสมอ
    git remote set-url origin %REPO_URL% >nul 2>&1
    if errorlevel 1 git remote add origin %REPO_URL%
    echo [OK] พบ Git repository
    echo.
)

:: ── แสดง remote ─────────────────────────────────────
echo [INFO] Remote URL:
git remote -v
echo.

:: ── Commit message ───────────────────────────────────
set /p COMMIT_MSG="ใส่ commit message (Enter = ใช้ default): "
if "%COMMIT_MSG%"=="" set COMMIT_MSG=Update dashboard data and code

:: ── Stage + Commit ───────────────────────────────────
echo.
echo [GIT] กำลัง add files...
git add .

echo.
echo [GIT] ไฟล์ที่เปลี่ยนแปลง:
git status --short
echo.

git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo [INFO] ไม่มีการเปลี่ยนแปลง — ไม่ต้อง push
    echo.
    pause
    exit /b 0
)

:: ── Push ─────────────────────────────────────────────
echo.
echo [GIT] กำลัง push ไปยัง GitHub...
git push -u origin main
if errorlevel 1 (
    echo.
    echo [ERROR] Push ไม่สำเร็จ
    echo.
    echo สาเหตุที่พบบ่อย:
    echo   1. ยังไม่ได้ login — ต้องใช้ Personal Access Token
    echo      สร้าง token ที่: https://github.com/settings/tokens
    echo      เลือก scope: repo
    echo.
    echo   2. ตอน Windows ถามให้ใส่ credentials:
    echo      Username: zterollout
    echo      Password: ^<วาง Personal Access Token^>
    echo.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo   [SUCCESS] Push สำเร็จ!
echo =====================================================
echo.
echo   Render.com จะ auto-deploy ภายใน 2-3 นาที
echo.
echo   ติดตาม deploy status:
echo   https://dashboard.render.com
echo.
echo   URL เว็บไซต์:
echo   %SITE_URL%
echo.
set /p OPEN_WEB="เปิด browser ดู deploy status? (Y/N): "
if /i "%OPEN_WEB%"=="Y" start "" "https://dashboard.render.com"
echo.
pause
