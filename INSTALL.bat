@echo off
setlocal
title Mutakamil Plus - Install
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel%==0 (set "PY=py") else (
  where python >nul 2>&1
  if %errorlevel%==0 (set "PY=python") else (
    echo Python was not found.
    echo Please install Python 3.11 or newer.
    pause
    exit /b 1
  )
)
echo Python:
%PY% --version
echo.
echo Installing packages...
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
  echo Installation failed.
  pause
  exit /b 1
)
%PY% -c "import openpyxl,PIL,pytesseract,fitz; print('Required Python packages OK')"
if errorlevel 1 (
  echo Package verification failed.
  pause
  exit /b 1
)
echo.
echo Installation completed.
echo PDF/image OCR additionally requires Tesseract OCR for Windows.
echo Excel/CSV matching is available without Tesseract.
pause
endlocal
