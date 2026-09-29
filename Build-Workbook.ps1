<#
.SYNOPSIS
    Builds M365_Audit_Dashboard.xlsm from M365_Audit_Dashboard.xlsx + the VBA
    modules, and binds the RUN AUDIT button. Run once on your Windows PC.

.DESCRIPTION
    Uses Excel COM automation to:
      1. open dist\M365_Audit_Dashboard.xlsx,
      2. import vba\modAudit.bas and vba\modImport.bas,
      3. draw a RUN AUDIT form-control button on the Dashboard sheet bound to
         the RunAudit macro,
      4. save as M365_Audit_Dashboard.xlsm next to the xlsx.

    ONE-TIME PREREQUISITE (Excel):
      File > Options > Trust Center > Trust Center Settings > Macro Settings >
      tick "Trust access to the VBA project object model".
    Without it, Excel forbids importing VBA code programmatically and this
    script stops with a clear message.

.EXAMPLE
    .\Build-Workbook.ps1
#>
[CmdletBinding()]
param(
    [string]$SourceWorkbook = "",
    [string]$Destination = ""
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($SourceWorkbook -eq '') { $SourceWorkbook = Join-Path $scriptDir 'M365_Audit_Dashboard.xlsx' }
if ($Destination -eq '')    { $Destination    = Join-Path $scriptDir 'M365_Audit_Dashboard.xlsm' }
$vbaDir = Join-Path $scriptDir 'vba'

if (-not (Test-Path $SourceWorkbook)) { throw "Source workbook not found: $SourceWorkbook" }
foreach ($m in @('modAudit.bas', 'modImport.bas')) {
    if (-not (Test-Path (Join-Path $vbaDir $m))) { throw "VBA module not found: $vbaDir\$m" }
}

Write-Host "Opening Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($SourceWorkbook)

    # VBA project access must be trusted (one-time Excel option)
    $vbProject = $null
    try { $vbProject = $wb.VBProject } catch { }
    if ($null -eq $vbProject) {
        throw ("Excel blocks access to the VBA project. One-time fix: Excel > File > Options > " +
               "Trust Center > Trust Center Settings > Macro Settings > tick " +
               "'Trust access to the VBA project object model', then run this script again.")
    }

    Write-Host "Importing VBA modules ..."
    foreach ($m in @('modAudit.bas', 'modImport.bas')) {
        # replace an existing module of the same name (idempotent re-runs)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($m)
        foreach ($comp in @($vbProject.VBComponents)) {
            if ($comp.Name -eq $name) { $vbProject.VBComponents.Remove($comp) }
        }
        $null = $vbProject.VBComponents.Import((Join-Path $vbaDir $m))
        Write-Host "  imported $m"
    }

    Write-Host "Adding the RUN AUDIT button ..."
    $dash = $wb.Worksheets.Item('Dashboard')
    foreach ($btn in @($dash.Buttons())) {
        if ($btn.Caption -like '*RUN AUDIT*') { $btn.Delete() }
    }
    # placed over S1:T2, next to the title banner
    $anchor = $dash.Range('S1')
    $btnObj = $dash.Buttons().Add($anchor.Left, $anchor.Top + 2, 95, 26)
    $btnObj.Caption = 'RUN AUDIT'
    $btnObj.OnAction = 'RunAudit'
    $btnObj.Font.Bold = $true

    Write-Host "Saving $Destination ..."
    if (Test-Path $Destination) { Remove-Item $Destination -Force }
    $wb.SaveAs($Destination, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
    $wb = $null
    Write-Host "Done: $Destination" -ForegroundColor Green
    Write-Host "Keep Collect-M365Audit.ps1 in the same folder as the .xlsm."
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
