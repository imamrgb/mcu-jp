@echo off
setlocal
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python belum terpasang. Gunakan URL Back4App atau instal Python terlebih dahulu.
  pause
  exit /b 1
)
python -c "import docx" >nul 2>nul
if errorlevel 1 (
  echo Menginstal dependency Word...
  python -m pip install -r requirements.txt
  if errorlevel 1 pause & exit /b 1
)
start "MCU Local Server" cmd /k "cd /d "%~dp0" && set PORT=8080 && python serve_railway.py"
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8080/index.html"
endlocal
