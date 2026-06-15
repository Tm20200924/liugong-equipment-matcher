@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo   LiuGong Equipment Matcher v19
echo ========================================
echo.
echo [1] Streamlit Web App (??)
echo [2] Command Line Mode
echo [3] Exit
echo.
set /p choice="Select (1/2/3): "

if "%choice%"=="1" (
    echo Starting Streamlit...
    streamlit run app.py
) else if "%choice%"=="2" (
    echo Running command-line engine...
    python engine/engine_v19.py
    pause
) else (
    exit
)
