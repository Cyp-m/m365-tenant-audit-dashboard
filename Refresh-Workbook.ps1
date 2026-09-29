<#
.SYNOPSIS
    Macro-free fallback importer: loads DATA\*.csv and MANUAL\*.csv into the
    workbook through Excel COM, for PCs where macros are blocked.

.DESCRIPTION
    The workbook must be CLOSED in Excel before running this script (the script
    opens its own hidden Excel instance, imports, recalculates, saves, closes).
    It performs exactly the same import as the VBA modImport module: same file
    names, same header checks, same explicit column types (comma delimiter,
    Local:=False), same RunInfo stamping.

.PARAMETER WorkbookPath
    Path to M365_Audit_Dashboard.xlsm (or .xlsx). Default: next to this script,
    .xlsm preferred when both exist.

.PARAMETER DataPath / ManualPath
    DATA and MANUAL folders. Default: next to the workbook.

.EXAMPLE
    .\Refresh-Workbook.ps1
#>
[CmdletBinding()]
param(
    [string]$WorkbookPath = "",
    [string]$DataPath = "",
    [string]$ManualPath = ""
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if ($WorkbookPath -eq '') {
    $xlsm = Join-Path $scriptDir 'M365_Audit_Dashboard.xlsm'
    $xlsx = Join-Path $scriptDir 'M365_Audit_Dashboard.xlsx'
    if (Test-Path $xlsm) { $WorkbookPath = $xlsm }
    elseif (Test-Path $xlsx) { $WorkbookPath = $xlsx }
    else { throw "No workbook found next to this script." }
}
$wbDir = Split-Path -Parent (Resolve-Path $WorkbookPath)
if ($DataPath -eq '') { $DataPath = Join-Path $wbDir 'DATA' }
if ($ManualPath -eq '') { $ManualPath = Join-Path $wbDir 'MANUAL' }
if (-not (Test-Path $DataPath)) { throw "DATA folder not found: $DataPath. Run Collect-M365Audit.ps1 first." }

# same contract as vba\modImport.bas: csv|sheet|max rows|types (T text, D date, N number, B bool)
$Datasets = @(
    'Users.csv|D_Users|20000|TTTTTBBTDNB',
    'Groups.csv|D_Groups|10000|TTTTTB',
    'MFA.csv|D_MFA|20000|TBBBT',
    'Roles.csv|D_Roles|4000|TTTTT',
    'EnterpriseApps.csv|D_EnterpriseApps|5000|TTTDBTB',
    'AppRegistrations.csv|D_AppRegistrations|3000|TTDNDN',
    'Licenses.csv|D_Licenses|300|TTNNN',
    'Mailboxes.csv|D_Mailboxes|20000|TTTNNDDDBBTTB',
    'MailboxPermissions.csv|D_MailboxPermissions|10000|TTT',
    'DistributionGroups.csv|D_DistributionGroups|5000|TTTN',
    'DLMembers.csv|D_DLMembers|30000|TTTT',
    'Domains.csv|D_Domains|300|TTBBN',
    'TransportRules.csv|D_TransportRules|1000|TTNT',
    'EmailActivity.csv|D_EmailActivity|400|DNNN',
    'EmailUserActivity.csv|D_EmailUserActivity|20000|TDNNN',
    'SPOSites.csv|D_SPOSites|10000|TTDNNNTB',
    'OneDrive.csv|D_OneDrive|20000|TTDNNN',
    'Teams.csv|D_Teams|10000|TTTNNND',
    'TeamsUserActivity.csv|D_TeamsUserActivity|20000|TDNNNN',
    'DevicesEntra.csv|D_DevicesEntra|20000|TTTTDBBD',
    'DevicesIntune.csv|D_DevicesIntune|20000|TTTTDTTTD',
    'SecuritySettings.csv|D_SecuritySettings|2|BNNTBB'
)
$ManualSets = @(
    'TeamsAdminExport.csv|M_TeamsAdminExport|10000|1',
    'SPAdminSites.csv|M_SPAdminSites|10000|1',
    'MDE_Devices.csv|M_MDE_Devices|20000|1',
    'ExternalFileAccess.csv|M_ExternalFileAccess|50000|1',
    'ITCosts.csv|M_ITCosts|1000|0'
)

function Set-CellValue {
    # PS 5.1 COM binding of chained "Cells.Item(r,c).Value2 = v" is unreliable
    # (misbinds to a String overload for some values); use an intermediate Range.
    param($Sheet, [int]$Row, [int]$Col, $Value)
    $rng = $Sheet.Cells.Item($Row, $Col)
    if ($Value -is [string]) { $rng.Value2 = [string]$Value }
    else { $rng.Value2 = [double]$Value }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rng) | Out-Null
}

function Get-JsonValue {
    param([string]$Json, [string]$Key)
    $m = [regex]::Match($Json, ('"{0}"\s*:\s*"((?:[^"\\]|\\.)*)"' -f [regex]::Escape($Key)))
    if ($m.Success) { return $m.Groups[1].Value.Replace('\"', '"') }
    $m = [regex]::Match($Json, ('"{0}"\s*:\s*([^,\r\n\}}\]]+)' -f [regex]::Escape($Key)))
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return ''
}

function Import-OneCsv {
    param($Excel, $Workbook, [string]$CsvFile, [string]$SheetName, [long]$MaxRows,
          [string]$Types, [bool]$AssertHeaders)
    $result = @{ rows = 0; status = 'OK'; message = '' }
    $dest = $Workbook.Worksheets.Item($SheetName)
    $nCols = 0
    if ($Types -ne '') {
        $nCols = $Types.Length
    } else {
        $nCols = $dest.Cells.Item(1, $dest.Columns.Count).End(-4159).Column   # xlToLeft
        while ($nCols -gt 1 -and "$($dest.Cells.Item(1, $nCols).Value2)" -like '*MANUAL INPUT*') { $nCols-- }
        while ($nCols -gt 1 -and "$($dest.Cells.Item(1, $nCols).Value2)".Trim() -eq '') { $nCols-- }
    }

    $fi = New-Object object[] $nCols
    for ($c = 1; $c -le $nCols; $c++) {
        $code = 1
        if ($Types.Length -ge $c) {
            switch ($Types.Substring($c - 1, 1)) {
                'T' { $code = 2 }
                'D' { $code = 5 }
                default { $code = 1 }
            }
        }
        $fi[$c - 1] = @($c, $code)
    }

    # OpenText(Filename, Origin, StartRow, DataType(1=xlDelimited), TextQualifier(1=double quote),
    #          ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo)
    # Trailing optional args are omitted: passing $null breaks COM late binding,
    # and the default Local=False is exactly what the contract requires.
    $Excel.Workbooks.OpenText($CsvFile, 65001, 1, 1, 1, $false, $false, $false, $true,
        $false, $false, [Type]::Missing, $fi)
    $csvWb = $Excel.ActiveWorkbook
    try {
        $csvWs = $csvWb.Worksheets.Item(1)
        if ($AssertHeaders) {
            for ($c = 1; $c -le $nCols; $c++) {
                $got = "$($csvWs.Cells.Item(1, $c).Value2)".Trim()
                $expected = "$($dest.Cells.Item(1, $c).Value2)".Trim()
                if ($got.ToLower() -ne $expected.ToLower()) {
                    $result.status = 'FAIL'
                    $result.message = "Header mismatch in $(Split-Path -Leaf $CsvFile) column ${c}: got '$got', expected '$expected'. File not imported."
                    return $result
                }
            }
        }
        $nRows = $csvWs.UsedRange.Rows.Count - 1
        if ($nRows -gt $MaxRows) {
            $result.status = 'WARN'
            $result.message = "File has $nRows rows; only the first $MaxRows were imported."
            $nRows = $MaxRows
        }
        # clear ONLY the contract columns so helper formula columns survive
        $dest.Range($dest.Cells.Item(2, 1), $dest.Cells.Item($MaxRows + 1, $nCols)).ClearContents() | Out-Null
        if ($nRows -gt 0) {
            $dest.Range('A2').Resize($nRows, $nCols).Value2 = $csvWs.Range('A2').Resize($nRows, $nCols).Value2
        }
        $result.rows = $nRows
        return $result
    } finally {
        $csvWb.Close($false)
    }
}

Write-Host "Opening $WorkbookPath (hidden Excel instance) ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: macros never run
$wb = $null
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $WorkbookPath).Path)
    $xl.Calculation = -4135   # xlCalculationManual

    $runInfo = $wb.Worksheets.Item('RunInfo')
    $runInfo.Range('A11:D51').ClearContents() | Out-Null
    $logRow = 11
    $ok = 0; $fail = 0; $missing = 0; $totalRows = 0

    # collector-side status
    $collStatus = @{}; $collMessage = @{}
    $statusFile = Join-Path $DataPath 'datasets_status.csv'
    if (Test-Path $statusFile) {
        foreach ($row in (Import-Csv $statusFile)) {
            $collStatus[$row.Dataset] = $row.Status
            $collMessage[$row.Dataset] = $row.Message
        }
    }

    foreach ($entry in $Datasets) {
        $p = $entry.Split('|')
        $csv = Join-Path $DataPath $p[0]
        $key = $p[0].Replace('.csv', '')
        if (-not (Test-Path $csv)) {
            $missing++
            Set-CellValue $runInfo $logRow 1 ([string]$key)
            Set-CellValue $runInfo $logRow 3 'MISSING'
            Set-CellValue $runInfo $logRow 4 "$($p[0]) not found in DATA folder"
            $logRow++
            continue
        }
        $r = $null
        try {
            $r = Import-OneCsv -Excel $xl -Workbook $wb -CsvFile $csv -SheetName $p[1] `
                -MaxRows ([long]$p[2]) -Types $p[3] -AssertHeaders $true
        } catch {
            $r = @{ rows = 0; status = 'FAIL'; message = $_.Exception.Message }
        }
        if ($r.status -eq 'OK' -and $collStatus.ContainsKey($key) -and $collStatus[$key] -ne 'OK') {
            $r.status = $collStatus[$key]; $r.message = $collMessage[$key]
        }
        if ($r.status -eq 'OK' -or $r.status -eq 'WARN') { $ok++ } else { $fail++ }
        $totalRows += $r.rows
        Write-Host ("  {0}: {1} ({2} rows) {3}" -f $key, $r.status, $r.rows, $r.message)
        Set-CellValue $runInfo $logRow 1 ([string]$key)
        Set-CellValue $runInfo $logRow 2 ([double]$r.rows)
        Set-CellValue $runInfo $logRow 3 ([string]$r.status)
        Set-CellValue $runInfo $logRow 4 ([string]$r.message)
        $logRow++
    }

    foreach ($entry in $ManualSets) {
        $p = $entry.Split('|')
        $csv = Join-Path $ManualPath $p[0]
        if (-not (Test-Path $csv)) { continue }
        $r = $null
        try {
            $r = Import-OneCsv -Excel $xl -Workbook $wb -CsvFile $csv -SheetName $p[1] `
                -MaxRows ([long]$p[2]) -Types '' -AssertHeaders ($p[3] -eq '1')
        } catch {
            $r = @{ rows = 0; status = 'FAIL'; message = $_.Exception.Message }
        }
        Write-Host ("  MANUAL {0}: {1} ({2} rows) {3}" -f $p[0], $r.status, $r.rows, $r.message)
        Set-CellValue $runInfo $logRow 1 ('MANUAL ' + $p[0].Replace('.csv', ''))
        Set-CellValue $runInfo $logRow 2 ([double]$r.rows)
        Set-CellValue $runInfo $logRow 3 ([string]$r.status)
        Set-CellValue $runInfo $logRow 4 ([string]$r.message)
        $logRow++
    }

    # RunInfo header from run.json
    $runJson = Join-Path $DataPath 'run.json'
    if (Test-Path $runJson) {
        $json = Get-Content $runJson -Raw -Encoding UTF8
        Set-CellValue $runInfo 3 2 ([string](Get-JsonValue $json 'tenantDisplayName'))
        Set-CellValue $runInfo 5 2 ([string](Get-JsonValue $json 'account'))
        Set-CellValue $runInfo 6 2 ([string](Get-JsonValue $json 'tenantId'))
        $concealed = 'No'
        if ((Get-JsonValue $json 'concealedNames') -match '(?i)true') { $concealed = 'Yes - per-user reports are pseudonymized' }
        Set-CellValue $runInfo 7 2 $concealed
        $wc = 0.0
        [void][double]::TryParse((Get-JsonValue $json 'warningsCount'), [ref]$wc)
        Set-CellValue $runInfo 8 2 $wc
        $iso = Get-JsonValue $json 'runStartUtc'
        if ($iso.Length -ge 10) {
            Set-CellValue $runInfo 4 2 ([datetime]::ParseExact($iso.Substring(0, 10), 'yyyy-MM-dd', $null).ToOADate())
        }
    }

    $xl.Calculation = -4105   # xlCalculationAutomatic
    $xl.CalculateFullRebuild()
    while ($xl.CalculationState -ne 0) { Start-Sleep -Milliseconds 200 }

    $wb.Save()
    $wb.Close($false)
    $wb = $null
    Write-Host ""
    Write-Host ("Refresh done. Datasets OK: {0}, failed: {1}, missing: {2}, rows: {3}." -f $ok, $fail, $missing, $totalRows) -ForegroundColor Green
    Write-Host "Open the workbook to see the dashboard."
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
