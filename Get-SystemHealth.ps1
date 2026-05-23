# Get-SystemHealth.ps1
# Basic IT infrastructure health check script

$hostname = $env:COMPUTERNAME
$timestamp = Get-Date

Write-Host "============================="
Write-Host "System Health Report"
Write-Host "Host: $hostname"
Write-Host "Time: $timestamp"
Write-Host "============================="

# CPU Load
$cpu = Get-CimInstance win32_processor | Measure-Object -Property LoadPercentage -Average
Write-Host "CPU Load: $($cpu.Average)%"

# RAM Usage
$os = Get-CimInstance Win32_OperatingSystem
$totalRAM = [math]::Round($os.TotalVisibleMemorySize/1MB,2)
$freeRAM = [math]::Round($os.FreePhysicalMemory/1MB,2)
$usedRAM = $totalRAM - $freeRAM

Write-Host "RAM Total: $totalRAM GB"
Write-Host "RAM Used: $usedRAM GB"
Write-Host "RAM Free: $freeRAM GB"

# Disk Usage
Write-Host "`nDisk Usage:"
Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
    $size = [math]::Round($_.Size/1GB,2)
    $free = [math]::Round($_.FreeSpace/1GB,2)
    $used = $size - $free
    $percent = [math]::Round(($used/$size)*100,2)

    Write-Host "$($_.DeviceID) - Used: $percent% ($used GB / $size GB)"
}

# Uptime
$uptime = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
$uptimeSpan = (Get-Date) - $uptime

Write-Host "`nUptime: $([math]::Round($uptimeSpan.TotalDays,2)) days"

Write-Host "============================="
