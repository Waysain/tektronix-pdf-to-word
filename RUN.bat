@echo off
setlocal
cd /d "%~dp0"
title PDF to Word Converter - 10BASE-T and 100BASE-TX
where py >nul 2>nul
if errorlevel 1 (
  echo Python Launcher was not found. Install Python 3.8 x64.
  pause
  exit /b 1
)
py -3.8 -m pip install --disable-pip-version-check "PyMuPDF==1.21.1" "python-docx==0.8.11"
if errorlevel 1 (
  echo Package installation failed. Check internet connection.
  pause
  exit /b 1
)
py -3.8 "%~dp0main.py"
if errorlevel 1 pause
endlocal
