<#
.SYNOPSIS
    Builds self-contained convert_csv_to_xlsx.exe next to the post (no .NET install required on target PC).

.USAGE
    powershell -ExecutionPolicy Bypass -File DEV_ONLY\build_converter.ps1
#>

$ErrorActionPreference = 'Stop'
$devRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$postRoot = Split-Path -Parent $devRoot
$proj = Join-Path $devRoot 'ConvertCsvToXlsx\ConvertCsvToXlsx.csproj'
$outDir = Join-Path $devRoot 'ConvertCsvToXlsx\bin\publish'

Write-Host 'Publishing convert_csv_to_xlsx.exe (ClosedXML 0.105, win-x64, self-contained)...' -ForegroundColor Cyan
dotnet publish $proj -c Release -o $outDir

$exe = Join-Path $outDir 'convert_csv_to_xlsx.exe'
$dest = Join-Path $postRoot 'convert_csv_to_xlsx.exe'
Copy-Item -Force $exe $dest
Write-Host "OK: $dest" -ForegroundColor Green
Write-Host 'Deploy convert_csv_to_xlsx.exe alongside shop_doc_advanced.tcl (post root folder).'
