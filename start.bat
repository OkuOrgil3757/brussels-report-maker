@echo off
echo ============================================================
echo   Brussels Report Maker
echo ============================================================
echo.

where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python not found. Please install Python 3.10+ from python.org
    pause
    exit /b 1
)

echo Checking dependencies...
pip install -r requirements.txt -q

echo.
echo Starting server...
echo Open http://localhost:8080 in your browser
echo Press Ctrl+C to stop.
echo.
python app.py
pause
