@echo off
setlocal
title Mutakamil Plus - Start
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel%==0 (set "PY=py") else (
  where python >nul 2>&1
  if %errorlevel%==0 (set "PY=python") else (
    echo Python was not found. Run INSTALL.bat first.
    pause
    exit /b 1
  )
)
%PY% app.py
pause
endlocal
