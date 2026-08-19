# check_bot.ps1 - Cek apakah bot_telegram.py sedang berjalan
$procs = Get-CimInstance Win32_Process | Where-Object {
    $_.CommandLine -match 'bot_telegram\.py' -and $_.Name -match 'python'
}
if ($procs) {
    exit 1  # Bot sedang berjalan
} else {
    exit 0  # Bot tidak berjalan
}
