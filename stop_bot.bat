@echo off
echo ===================================================
echo     MENGHENTIKAN BOT TELEGRAM - bot_telegram.py
echo ===================================================
echo.
echo Mencari proses bot...

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0stop_bot_action.ps1"
if %errorlevel%==1 (
    echo [INFO] Tidak ada program Bot Telegram yang sedang berjalan.
) else (
    echo [OK] Bot berhasil dimatikan dari background!
)

echo.
echo ===================================================
pause
