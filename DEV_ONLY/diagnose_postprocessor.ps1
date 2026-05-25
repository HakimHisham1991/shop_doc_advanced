<#
.SYNOPSIS
    Snapshot stuck processes, temp files, and timing hints for shop_doc post slowdown.

.DESCRIPTION
    Run this AFTER you notice the NX post dialog getting slow (before killing anything).
    Safe for guest users — read-only checks except optional -KillExcelZombies.

.USAGE
    powershell -ExecutionPolicy Bypass -File DEV_ONLY\diagnose_postprocessor.ps1
    powershell -ExecutionPolicy Bypass -File DEV_ONLY\diagnose_postprocessor.ps1 -KillExcelZombies
#>

param(
    [switch]$KillExcelZombies
)

$ErrorActionPreference = 'SilentlyContinue'
$stamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$report = Join-Path $env:TEMP "shopdoc_diagnose_$stamp.txt"

function Write-Report($text) {
    $text | Tee-Object -FilePath $report -Append
}

Write-Report "=== Shop Doc postprocessor diagnostic $stamp ==="
Write-Report "User: $env:USERNAME  Computer: $env:COMPUTERNAME"
Write-Report ""

Write-Report "--- Processes (Excel / cscript / convert_csv / NX) ---"
Get-Process | Where-Object {
    $_.ProcessName -match '^(EXCEL|cscript|wscript|convert_csv_to_xlsx|ugraf|nx|mom)$'
} | Sort-Object ProcessName, StartTime |
    Format-Table ProcessName, Id, StartTime, @{N='WorkingSetMB';E={[math]::Round($_.WorkingSet64/1MB,1)}} -AutoSize |
    Out-String | ForEach-Object { Write-Report $_ }

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
    Write-Report "WARNING: $($excel.Count) EXCEL.EXE instance(s) — classic COM orphan cause of gradual slowdown."
    if ($KillExcelZombies) {
        $excel | Stop-Process -Force
        Write-Report "Killed EXCEL.EXE processes."
    }
} else {
    Write-Report "OK: No EXCEL.EXE processes."
}

$cscript = Get-Process cscript -ErrorAction SilentlyContinue
if ($cscript) {
    Write-Report "WARNING: $($cscript.Count) cscript.exe instance(s) still running."
}

Write-Report ""
Write-Report "--- TEMP mom_pause scratch files ---"
Get-ChildItem $env:TEMP -Filter '*_mom_pause_*.txt' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 20 Name, Length, LastWriteTime |
    Format-Table -AutoSize |
    Out-String | ForEach-Object { Write-Report $_ }

Write-Report ""
Write-Report "--- Shop doc launcher / probe temp files in TEMP ---"
Get-ChildItem $env:TEMP -Filter '_shopdoc*' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 20 FullName, LastWriteTime |
    Format-Table -AutoSize |
    Out-String | ForEach-Object { Write-Report $_ }

Write-Report ""
Write-Report "--- Converter next to post (if UGII_CAM_POST_DIR or post root) ---"
$devRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$postRoot = Split-Path -Parent $devRoot
$postDirs = @(
    $postRoot,
    $env:UGII_CAM_POST_DIR
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

foreach ($dir in $postDirs) {
    Write-Report "Post dir: $dir"
    Get-ChildItem $dir -Filter 'convert_csv_to_xlsx*' -ErrorAction SilentlyContinue |
        Select-Object Name, Length, LastWriteTime |
        Format-Table -AutoSize |
        Out-String | ForEach-Object { Write-Report $_ }
}

Write-Report ""
Write-Report "--- Quick Excel COM registration (informational only) ---"
foreach ($key in @('Excel.Application', 'Excel.Application.16')) {
    $out = reg query "HKCR\$key" /ve 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Report "Registered: $key"
    } else {
        Write-Report "Not registered: $key"
    }
}

Write-Report ""
Write-Report "Report saved: $report"
Write-Report "If EXCEL count grows after each post, migrate to convert_csv_to_xlsx.exe (ClosedXML, no COM)."
