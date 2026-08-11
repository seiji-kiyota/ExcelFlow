@echo off
cd /d "%~dp0"

echo Starting ExcelFlow...
echo URL: http://localhost:8501

REM Open the browser a few seconds after Streamlit begins starting.
start "" cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:8501"

streamlit run app.py --server.headless true
pause
