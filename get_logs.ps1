# get_logs.ps1
$TimeFilter = (Get-Date).AddDays(-1) # ย้อนหลัง 1 วัน

# ดึง Error จาก System และ Application Logs
try {
    $Logs = Get-WinEvent -FilterHashtable @{
        LogName = 'System','Application'; 
        Level = 2; # 2 คือ Error เท่านั้น
        StartTime = $TimeFilter
    } -ErrorAction SilentlyContinue | Select-Object TimeCreated, LogName, Id, Message

    # เซฟเป็น CSV เพื่อให้ Python อ่านง่าย
    $Logs | Export-Csv -Path "current_logs.csv" -NoTypeInformation -Encoding UTF8
    Write-Host "Success: ข้อมูลถูกเขียนลงไฟล์ current_logs.csv แล้ว"
}
catch {
    Write-Host "Error: ไม่สามารถดึงข้อมูลได้"
}