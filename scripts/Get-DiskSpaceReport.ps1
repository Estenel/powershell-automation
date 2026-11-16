<# 
.SYNOPSIS
Reports free space on local fixed drives.

.DESCRIPTION
Shows each local fixed drive's total size, free space, and free-space percentage.
Drives below WarningFreePercent are marked as warnings. The script only reads
drive information and does not change files or disk settings.

.PARAMETER WarningFreePercent
Mark a drive as a warning when its free-space percentage is below this value.
Defaults to 15.

.EXAMPLE
.\Get-DiskSpaceReport.ps1

.EXAMPLE
.\Get-DiskSpaceReport.ps1 -WarningFreePercent 10
#>
[CmdletBinding()]
param(
    [ValidateRange(0, 100)]
    [int] $WarningFreePercent = 15
)

Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 3' |
    Where-Object { $_.Size -gt 0 } |
    ForEach-Object {
        $freePercent = [math]::Round(($_.FreeSpace / $_.Size) * 100, 1)
        [pscustomobject]@{
            Drive             = $_.DeviceID
            VolumeName        = $_.VolumeName
            SizeGB            = [math]::Round($_.Size / 1GB, 2)
            FreeGB            = [math]::Round($_.FreeSpace / 1GB, 2)
            FreePercent       = $freePercent
            Status            = if ($freePercent -lt $WarningFreePercent) { 'Warning' } else { 'OK' }
        }
    } |
    Sort-Object Drive
