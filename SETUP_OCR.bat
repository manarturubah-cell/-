@echo off
setlocal
cd /d "%~dp0"
title Mutakamil Plus - OCR Setup

echo ================================================
echo  إعداد OCR للفواتير PDF والصور
 echo ================================================
echo.

where winget >nul 2>&1
if errorlevel 1 (
  echo لم يتم العثور على winget في هذا الجهاز.
  echo.
  echo افتح صفحة التثبيت الرسمية لـ Tesseract على GitHub:
  echo https://github.com/UB-Mannheim/tesseract/wiki
  echo وثبّت Tesseract مع لغة Arabic / ara.
  echo.
  pause
  exit /b 1
)

echo سيتم تثبيت Tesseract OCR عبر Windows Package Manager.
echo يجب اختيار/تثبيت لغة العربية عند توفر خيار اللغات.
echo.
winget install -e --id UB-Mannheim.TesseractOCR
if errorlevel 1 (
  echo.
  echo تعذر التثبيت التلقائي.
  echo يمكنك تثبيته يدويًا من:
  echo https://github.com/UB-Mannheim/tesseract/wiki
  pause
  exit /b 1
)

echo.
echo ================================================
echo تم تثبيت Tesseract. أغلق التطبيق وأعد تشغيل START.bat.
echo ================================================
pause
endlocal
