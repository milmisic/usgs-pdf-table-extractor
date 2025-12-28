@echo off
echo ========================================
echo   USGS Table Extractor
echo ========================================
echo.

cd /d "%~dp0"
python run_extraction.py

echo.
echo ========================================
echo   Processing Complete!
echo ========================================
echo.
pause