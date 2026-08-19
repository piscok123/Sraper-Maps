# stop_bot_action.ps1 - Cari dan matikan proses bot_telegram.py
$procs = Get-CimInstance Win32_Process | Where-Object {
    $_.CommandLine -match 'bot_telegram\.py' -and $_.Name -match 'python'
}
if ($procs) {
    $procs | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
    exit 0  # Berhasil dimatikan
} else {
    exit 1  # Tidak ada proses
}
