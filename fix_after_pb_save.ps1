<#
.SYNOPSIS
    Patches shop_doc_advanced.tcl after a Post Builder save to re-insert
    the mom_stepover_distance trace setup into MOM_start_of_program.

.DESCRIPTION
    Post Builder regenerates MOM_start_of_program and strips any manual
    code. This script finds the stable anchor pattern and inserts the
    two trace lines that arm pb__trace_stepover before MOM_machine_mode
    loads operation parameters.

    Safe to run multiple times (idempotent).

.USAGE
    Right-click -> Run with PowerShell, or from a terminal:
        powershell -ExecutionPolicy Bypass -File fix_after_pb_save.ps1
#>

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$tclFile   = Join-Path $scriptDir 'shop_doc_advanced.tcl'

if (-not (Test-Path $tclFile)) {
    Write-Host 'ERROR: shop_doc_advanced.tcl not found.' -ForegroundColor Red
    Write-Host 'Place this script in the same folder as shop_doc_advanced.tcl'
    pause
    exit 1
}

$content = Get-Content $tclFile -Raw

$anchor  = '    rename PB_load_alternate_unit_settings ""'
$marker  = '#************'
$traceLn = 'trace add variable ::mom_stepover_distance write pb__trace_stepover'

$anchorPos = $content.IndexOf($anchor)
if ($anchorPos -lt 0) {
    Write-Host 'ERROR: Anchor pattern not found in TCL file.' -ForegroundColor Red
    pause
    exit 1
}

$markerPos = $content.IndexOf($marker, $anchorPos)
if ($markerPos -lt 0) {
    Write-Host 'ERROR: uplevel marker not found after anchor.' -ForegroundColor Red
    pause
    exit 1
}

# Check if trace line already exists between the anchor and the uplevel marker
$between = $content.Substring($anchorPos, $markerPos - $anchorPos)
if ($between.Contains($traceLn)) {
    Write-Host 'OK: Trace lines already present in MOM_start_of_program -- no changes needed.' -ForegroundColor Green
    pause
    exit 0
}

$patch = $anchor + "`r`n`r`n    catch { trace remove variable ::mom_stepover_distance write pb__trace_stepover }`r`n    trace add variable ::mom_stepover_distance write pb__trace_stepover`r`n"
$content = $content.Remove($anchorPos, $anchor.Length).Insert($anchorPos, $patch)

Set-Content $tclFile -Value $content -NoNewline
Write-Host 'PATCHED: Trace lines inserted into MOM_start_of_program.' -ForegroundColor Green
pause
