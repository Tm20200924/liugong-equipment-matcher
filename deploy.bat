@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==============================================
echo   LiuGong Matcher - GitHub Push & Deploy
echo ==============================================
echo.

REM Find gh
for /f "delims=" %%i in ('where /r "%LOCALAPPDATA%\ghcli" gh.exe 2^>nul') do set GH=%%i
if not defined GH set GH=gh

REM Check auth
%GH% auth status >nul 2>&1
if errorlevel 1 (
    echo [!] GitHub not authenticated.
    echo.
    echo Please run: %GH% auth login
    echo Or paste your token to work\.ghtoken and re-run.
    pause
    exit /b 1
)

echo [1/3] Creating GitHub repository...
%GH% repo create liugong-equipment-matcher --public --source=. --remote=origin --push -d "LiuGong Equipment Matching Engine - DAP price calculation with cross-verification"
if errorlevel 1 (
    echo [!] Repo creation failed. Trying push to existing...
    git remote add origin https://github.com/YOUR_USERNAME/liugong-equipment-matcher.git 2>nul
    git push -u origin master
)

echo.
echo [2/3] Pushing to GitHub...
git push -u origin master

echo.
echo [3/3] Done!
echo.
echo ==============================================
echo   Next: Deploy on Streamlit Cloud
echo ==============================================
echo   1. Open https://share.streamlit.io
echo   2. Sign in with GitHub
echo   3. Click "New app"
echo   4. Select repo: liugong-equipment-matcher
echo   5. Main file path: app.py
echo   6. Click Deploy!
echo ==============================================
pause
