@echo off
cd /d "%~dp0"

REM === Cek apakah bot sudah berjalan ===
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check_bot.ps1"
if %errorlevel%==1 (
    echo ===================================================
    echo  WARNING: BOT TELEGRAM SUDAH BERJALAN!
    echo ===================================================
    echo  Program bot_telegram.py sudah aktif di sistem.
    echo  Mencegah eksekusi ganda - double execution.
    echo ===================================================
    echo.
    pause
    exit /b 1
)

echo ===================================================
echo     MEMULAI BOT TELEGRAM - bot_telegram.py
echo ===================================================
call venv\Scripts\activate.bat
python bot_telegram.py
