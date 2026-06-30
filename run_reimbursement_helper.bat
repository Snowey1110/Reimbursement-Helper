@echo off
setlocal
cd /d "%~dp0"
python -c "import openpyxl, PIL, fitz" >nul 2>nul
if errorlevel 1 (
  echo Installing required packages...
  python -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo Could not install required packages. Please install Python and try again.
    pause
    exit /b 1
  )
)
python reimbursement_helper.py
if errorlevel 1 pause
endlocal
