<#
.SYNOPSIS
    M365 audit collector - read-only, Global Reader + delegated Graph consent.

.DESCRIPTION
    Signs in interactively to the target tenant (browser window), collects all
    audit datasets through Microsoft Graph and Exchange Online, and writes them
    as fixed-name, fixed-schema CSV files into the DATA folder next to the
    workbook. The workbook importer (RUN AUDIT button or Refresh-Workbook.ps1)
    then loads them and the dashboard recalculates itself.

    The tool NEVER writes anything to the tenant. Every call is a read.

.PARAMETER OutputPath
    Folder for the DATA CSVs. Default: <script folder>\DATA.

.PARAMETER SkipExchange
    Skip all Exchange Online datasets (Graph only).

.PARAMETER SkipGraph
    Skip all Microsoft Graph datasets (Exchange only).

.PARAMETER DeepMailboxScan
    Also read Inbox / Sent Items folder statistics per mailbox to get the exact
    LastEmailReceived / LastEmailSent dates (slow: 2 extra calls per mailbox,
    throttle-aware). Without it, those columns are estimated from the tenant
    email activity report (one call for the whole tenant).

.PARAMETER SampleData
    Offline test mode: point to reference/sample_data. No connection is made;
    the sample exports are transformed into the exact DATA\*.csv contract.

.PARAMETER IncludeExternalFileAccess
    Optional module: external user file access report (Search-UnifiedAuditLog).
    Needs an audit role (e.g. View-Only Audit Logs) that Global Reader does NOT
    have. The collector tests access first and skips with a warning if denied.

.PARAMETER UserPrincipalName
    Account used for the Exchange Online connection (skips the account picker).

.EXAMPLE
    .\Collect-M365Audit.ps1
    .\Collect-M365Audit.ps1 -OutputPath "C:\Audit\DATA" -DeepMailboxScan
    .\Collect-M365Audit.ps1 -SampleData ..\reference\sample_data
#>
[CmdletBinding()]
param(
    [string]$OutputPath = "",
    [switch]$SkipExchange,
    [switch]$SkipGraph,
    [switch]$DeepMailboxScan,
    [string]$SampleData = "",
    [switch]$IncludeExternalFileAccess,
    [string]$UserPrincipalName = "",
    [int]$ExternalAccessDays = 90
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Off

# ---------------------------------------------------------------------------
# Paths, logging
# ---------------------------------------------------------------------------
$script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $script:ScriptDir 'DATA' }
$script:DataPath   = $OutputPath
$script:ManualPath = Join-Path (Split-Path -Parent $script:DataPath) 'MANUAL'
$script:LogDir     = Join-Path (Split-Path -Parent $script:DataPath) 'logs'
foreach ($d in @($script:DataPath, $script:LogDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$script:LogFile = Join-Path $script:LogDir ("collector_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$script:Warnings      = New-Object System.Collections.Generic.List[string]
$script:DatasetStatus = [ordered]@{}
$script:RunStart      = (Get-Date).ToUniversalTime()

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Add-Warning {
    param([string]$Message)
    $script:Warnings.Add($Message) | Out-Null
    Write-Log $Message 'WARN'
}

# ---------------------------------------------------------------------------
# Formatting helpers (culture-proof: CSV must be identical on FR/EN machines)
# ---------------------------------------------------------------------------
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture

function Format-IsoDate {
    param($Value)
    if ($null -eq $Value -or $Value -eq '') { return '' }
    if ($Value -is [datetime]) {
        if ($Value.Year -lt 1971) { return '' }
        return $Value.ToString('yyyy-MM-dd', $script:Inv)
    }
    $s = [string]$Value
    if ($s -match '^\d{4}-\d{2}-\d{2}') { return $s.Substring(0, 10) }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($s, $script:Inv, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        if ($dt.Year -lt 1971) { return '' }
        return $dt.ToString('yyyy-MM-dd', $script:Inv)
    }
    return ''
}

function Format-Bool {
    param($Value, [string]$Default = 'FALSE')
    if ($null -eq $Value -or "$Value" -eq '') { return $Default }
    if ($Value -is [bool]) { if ($Value) { return 'TRUE' } else { return 'FALSE' } }
    $s = ([string]$Value).Trim()
    if ($s -match '^(true|yes|1|enabled|on)$') { return 'TRUE' }
    if ($s -match '^(false|no|0|disabled|off)$') { return 'FALSE' }
    return $Default
}

function Format-Num {
    param($Value, [int]$Decimals = -1)
    if ($null -eq $Value -or "$Value" -eq '') { return '' }
    $d = 0.0
    if (-not [double]::TryParse(([string]$Value), [System.Globalization.NumberStyles]::Any, $script:Inv, [ref]$d)) {
        if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) { $d = [double]$Value } else { return '' }
    }
    if ($Decimals -ge 0) { $d = [math]::Round($d, $Decimals) }
    return $d.ToString($script:Inv)
}

# CSV writer: UTF-8 with BOM, comma delimiter, every field quoted.
function Write-AuditCsv {
    param(
        [string]$FileName,
        [string[]]$Columns,
        $Rows,
        [string]$Folder = ""
    )
    if ($Folder -eq "") { $Folder = $script:DataPath }
    if (-not (Test-Path $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    $path = Join-Path $Folder $FileName
    $enc = New-Object System.Text.UTF8Encoding($true)
    $sw = New-Object System.IO.StreamWriter($path, $false, $enc)
    try {
        $header = @()
        foreach ($c in $Columns) { $header += ('"' + $c.Replace('"', '""') + '"') }
        $sw.WriteLine(($header -join ','))
        if ($null -ne $Rows) {
            foreach ($row in $Rows) {
                $vals = @()
                foreach ($c in $Columns) {
                    $v = $row[$c]
                    if ($null -eq $v) { $v = '' }
                    $s = ([string]$v).Replace('"', '""')
                    $s = $s -replace "(\r\n|\r|\n)", ' '
                    $vals += ('"' + $s + '"')
                }
                $sw.WriteLine(($vals -join ','))
            }
        }
    } finally { $sw.Close() }
}

# ---------------------------------------------------------------------------
# Dataset contract (must match the workbook D_* sheets - do not reorder)
# ---------------------------------------------------------------------------
$script:Contract = [ordered]@{
    'Users'              = @('Id','DisplayName','UserPrincipalName','Mail','UserType','AccountEnabled','OnPremSynced','UsageLocation','CreatedDateTime','LicenseCount','IsLicensed')
    'Groups'             = @('Id','DisplayName','Mail','Category','Visibility','IsTeam')
    'MFA'                = @('UserPrincipalName','IsMfaRegistered','IsMfaCapable','IsAdmin','MethodsRegistered')
    'Roles'              = @('RoleName','MemberDisplayName','MemberUpn','MemberType','AssignmentType')
    'EnterpriseApps'     = @('DisplayName','AppId','PublisherName','CreatedDateTime','AccountEnabled','Homepage','IsMicrosoftFirstParty')
    'AppRegistrations'   = @('DisplayName','AppId','CreatedDateTime','SecretCount','NearestSecretExpiry','CertCount')
    'Licenses'           = @('SkuPartNumber','FriendlyName','Total','Assigned','Available')
    'Mailboxes'          = @('DisplayName','PrimarySmtpAddress','RecipientTypeDetails','SizeGB','ItemCount','LastUserActionTime','LastEmailReceived','LastEmailSent','ArchiveEnabled','LitigationHold','ForwardingSmtpAddress','ForwardingAddress','ForwardsExternally')
    'MailboxPermissions' = @('Mailbox','PermissionType','Grantee')
    'DistributionGroups' = @('DisplayName','PrimarySmtpAddress','Type','MemberCountDirect')
    'DLMembers'          = @('DLName','MemberDisplayName','MemberSmtp','MemberRecipientType')
    'Domains'            = @('DomainName','Type','IsDefault','DkimEnabled','PasswordValidityDays')
    'TransportRules'     = @('Name','State','Priority','Comments')
    'EmailActivity'      = @('ReportDate','Send','Receive','Read')
    'EmailUserActivity'  = @('UserPrincipalName','LastActivityDate','SendCount','ReceiveCount','ReadCount')
    'SPOSites'           = @('SiteUrl','OwnerDisplayName','LastActivityDate','FileCount','ActiveFileCount','StorageUsedGB','RootWebTemplate','IsTeamsConnected')
    'OneDrive'           = @('OwnerUpn','OwnerDisplayName','LastActivityDate','FileCount','ActiveFileCount','StorageUsedGB')
    'Teams'              = @('TeamName','GroupId','Visibility','MemberCount','ActiveUsers90d','ChannelMessages90d','LastActivityDate')
    'TeamsUserActivity'  = @('UserPrincipalName','LastActivityDate','TeamChatMessages','PrivateChatMessages','Calls','Meetings')
    'DevicesEntra'       = @('DisplayName','OS','OSVersion','TrustType','LastSignIn','IsCompliant','IsManaged','RegisteredDateTime')
    'DevicesIntune'      = @('DeviceName','OS','OSVersion','ComplianceState','LastSyncDateTime','ManagementAgent','Manufacturer','Model','EnrolledDateTime')
    'SecuritySettings'   = @('SecurityDefaultsEnabled','CAPoliciesTotal','CAPoliciesEnabled','TenantSharingCapability','SmtpAuthDisabled','AuditEnabled')
}

function Invoke-Dataset {
    param([string]$Key, [scriptblock]$Script)
    $cols = $script:Contract[$Key]
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $status = 'OK'; $msg = ''; $rows = @()
    Write-Log "Collecting $Key ..."
    try {
        $result = & $Script
        if ($null -eq $result) { $rows = @() } else { $rows = @($result) }
    } catch {
        $status = 'FAIL'; $msg = $_.Exception.Message
        Add-Warning ("{0}: {1}" -f $Key, $msg)
        $rows = @()
    }
    try {
        Write-AuditCsv -FileName ($Key + '.csv') -Columns $cols -Rows $rows
    } catch {
        $status = 'FAIL'; $msg = "CSV write failed: " + $_.Exception.Message
        Add-Warning ("{0}: {1}" -f $Key, $msg)
    }
    $sw.Stop()
    $script:DatasetStatus[$Key] = [ordered]@{
        rows = $rows.Count; status = $status; message = $msg
        seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
    }
    Write-Log ("{0}: {1}, {2} rows, {3}s" -f $Key, $status, $rows.Count, $script:DatasetStatus[$Key].seconds)
}

# ---------------------------------------------------------------------------
# SKU friendly names (subset of Microsoft "Product names and service plan
# identifiers for licensing"). Must match SKU_MAP in tools/build_workbook.py.
# ---------------------------------------------------------------------------
$script:SkuFriendlyNames = @{
    'SPE_E3' = 'Microsoft 365 E3'
    'SPE_E5' = 'Microsoft 365 E5'
    'SPE_F1' = 'Microsoft 365 F3'
    'M365_F1' = 'Microsoft 365 F1'
    'SPB' = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
    'O365_BUSINESS_PREMIUM' = 'Microsoft 365 Business Standard'
    'SMB_BUSINESS' = 'Microsoft 365 Apps for business'
    'O365_BUSINESS' = 'Microsoft 365 Apps for business'
    'OFFICESUBSCRIPTION' = 'Microsoft 365 Apps for enterprise'
    'STANDARDPACK' = 'Office 365 E1'
    'ENTERPRISEPACK' = 'Office 365 E3'
    'ENTERPRISEPREMIUM' = 'Office 365 E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF' = 'Office 365 E5 without Audio Conferencing'
    'DESKLESSPACK' = 'Office 365 F3'
    'EXCHANGESTANDARD' = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE' = 'Exchange Online (Plan 2)'
    'EXCHANGEDESKLESS' = 'Exchange Online Kiosk'
    'EXCHANGEARCHIVE_ADDON' = 'Exchange Online Archiving'
    'SHAREPOINTSTANDARD' = 'SharePoint Online (Plan 1)'
    'SHAREPOINTENTERPRISE' = 'SharePoint Online (Plan 2)'
    'MCOSTANDARD' = 'Skype for Business Online (Plan 2)'
    'MCOEV' = 'Microsoft Teams Phone Standard'
    'MCOMEETADV' = 'Microsoft 365 Audio Conferencing'
    'MCOPSTN1' = 'Microsoft Teams Domestic Calling Plan'
    'MCOPSTN2' = 'Microsoft Teams Domestic and International Calling Plan'
    'PHONESYSTEM_VIRTUALUSER' = 'Microsoft Teams Phone Resource Account'
    'MEETING_ROOM' = 'Microsoft Teams Rooms Standard'
    'Microsoft_Teams_Rooms_Pro' = 'Microsoft Teams Rooms Pro'
    'TEAMS_ESSENTIALS_AAD' = 'Microsoft Teams Essentials'
    'TEAMS_EXPLORATORY' = 'Microsoft Teams Exploratory'
    'AAD_PREMIUM' = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2' = 'Microsoft Entra ID P2'
    'EMS' = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM' = 'Enterprise Mobility + Security E5'
    'INTUNE_A' = 'Microsoft Intune Plan 1'
    'ATP_ENTERPRISE' = 'Microsoft Defender for Office 365 (Plan 1)'
    'THREAT_INTELLIGENCE' = 'Microsoft Defender for Office 365 (Plan 2)'
    'DEFENDER_ENDPOINT_P1' = 'Microsoft Defender for Endpoint P1'
    'WIN_DEF_ATP' = 'Microsoft Defender for Endpoint P2'
    'ADALLOM_STANDALONE' = 'Microsoft Defender for Cloud Apps'
    'ATA' = 'Microsoft Defender for Identity'
    'RIGHTSMANAGEMENT' = 'Azure Information Protection Premium P1'
    'POWER_BI_STANDARD' = 'Power BI (free)'
    'POWER_BI_PRO' = 'Power BI Pro'
    'PBI_PREMIUM_PER_USER' = 'Power BI Premium Per User'
    'FLOW_FREE' = 'Power Automate Free'
    'POWERAUTOMATE_ATTENDED_RPA' = 'Power Automate Premium'
    'POWERAPPS_PER_USER' = 'Power Apps Premium'
    'PROJECT_P1' = 'Project Plan 1'
    'PROJECTPROFESSIONAL' = 'Project Plan 3'
    'PROJECTPREMIUM' = 'Project Plan 5'
    'VISIO_PLAN1_DEPT' = 'Visio Plan 1'
    'VISIOCLIENT' = 'Visio Plan 2'
    'WIN10_PRO_ENT_SUB' = 'Windows 10/11 Enterprise E3'
    'WIN10_VDA_E5' = 'Windows 10/11 Enterprise E5'
    'WINDOWS_STORE' = 'Windows Store for Business'
    'STREAM' = 'Microsoft Stream'
    'Microsoft_365_Copilot' = 'Microsoft 365 Copilot'
    'DEVELOPERPACK_E5' = 'Microsoft 365 E5 Developer'
    'FORMS_PRO' = 'Dynamics 365 Customer Voice Trial'
    'DYN365_ENTERPRISE_SALES' = 'Dynamics 365 Sales Enterprise'
    'DYN365_ENTERPRISE_CUSTOMER_SERVICE' = 'Dynamics 365 Customer Service Enterprise'
    'CCIBOTS_PRIVPREV_VIRAL' = 'Copilot Studio Viral Trial'
}

function Get-SkuFriendlyName {
    param([string]$SkuPartNumber)
    if ($script:SkuFriendlyNames.ContainsKey($SkuPartNumber)) {
        return $script:SkuFriendlyNames[$SkuPartNumber]
    }
    return $SkuPartNumber
}

# ---------------------------------------------------------------------------
# run.json + status CSV + done flag
# ---------------------------------------------------------------------------
function Write-RunMeta {
    param([string]$Outcome)
    $meta = [ordered]@{
        tenantId          = $script:TenantId
        tenantDisplayName = $script:TenantName
        account           = $script:Account
        runStartUtc       = $script:RunStart.ToString('yyyy-MM-ddTHH:mm:ssZ')
        runEndUtc         = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        outcome           = $Outcome
        concealedNames    = [bool]$script:ConcealedNames
        rolesOfAccount    = @($script:AccountRoles)
        warningsCount     = $script:Warnings.Count
        warnings          = @($script:Warnings)
        datasets          = $script:DatasetStatus
    }
    $json = $meta | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText((Join-Path $script:DataPath 'run.json'), $json, (New-Object System.Text.UTF8Encoding($true)))

    $statusRows = New-Object System.Collections.Generic.List[object]
    foreach ($key in $script:Contract.Keys) {
        if ($script:DatasetStatus.Contains($key)) {
            $d = $script:DatasetStatus[$key]
            $statusRows.Add([ordered]@{ Dataset = $key; Rows = $d.rows; Status = $d.status; Message = $d.message; Seconds = $d.seconds })
        } else {
            $statusRows.Add([ordered]@{ Dataset = $key; Rows = ''; Status = 'SKIPPED'; Message = 'not collected in this run'; Seconds = '' })
        }
    }
    Write-AuditCsv -FileName 'datasets_status.csv' -Columns @('Dataset','Rows','Status','Message','Seconds') -Rows $statusRows

    [System.IO.File]::WriteAllText(
        (Join-Path $script:DataPath '_DONE.flag'),
        ("{0} {1}" -f $Outcome, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
        (New-Object System.Text.UTF8Encoding($true)))
    Write-Log "Run finished: $Outcome. Warnings: $($script:Warnings.Count)"
}

$script:TenantId = ''; $script:TenantName = ''; $script:Account = ''
$script:ConcealedNames = $false; $script:AccountRoles = @()

# ===========================================================================
# SAMPLE DATA MODE (offline test harness) - no tenant connection at all
# ===========================================================================
function Convert-SampleData {
    param([string]$SamplePath)
    Write-Log "SAMPLE DATA MODE: transforming '$SamplePath' (no tenant connection)"
    if (-not (Test-Path $SamplePath)) { throw "Sample data path not found: $SamplePath" }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "Installing ImportExcel module (CurrentUser) ..."
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module ImportExcel

    function Find-Sample { param([string]$Pattern)
        $f = Get-ChildItem -Path $SamplePath -Recurse -File | Where-Object { $_.Name -like $Pattern } | Select-Object -First 1
        if ($null -eq $f) { throw "sample file not found: $Pattern" }
        return $f.FullName
    }
    function Get-Trimmed { param($Row, [string]$Name)
        foreach ($p in $Row.PSObject.Properties) {
            if ($p.Name.Trim() -eq $Name) { return $p.Value }
        }
        return $null
    }
    function Convert-LegacyDate { param($Value)
        # MBXStats mixes real datetimes, "dd/MM/yyyy" strings and "No*" sentinels
        if ($null -eq $Value) { return '' }
        if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd', $script:Inv) }
        $s = ([string]$Value).Trim()
        if ($s -eq '' -or $s -like 'No*' -or $s -eq 'N/A') { return '' }
        $dt = [datetime]::MinValue
        $fr = [System.Globalization.CultureInfo]::GetCultureInfo('fr-FR')
        if ([datetime]::TryParseExact($s, 'dd/MM/yyyy', $fr, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt.ToString('yyyy-MM-dd', $script:Inv)
        }
        return (Format-IsoDate $s)
    }

    # ---- Users --------------------------------------------------------------
    Invoke-Dataset 'Users' {
        $src = Import-Csv (Find-Sample 'exportUsers_*.csv')
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($u in $src) {
            $licCount = ([regex]::Matches([string]$u.assignedLicenses, '"skuId"')).Count
            $out.Add([ordered]@{
                Id = $u.id; DisplayName = $u.displayName; UserPrincipalName = $u.userPrincipalName
                Mail = $u.mail; UserType = $(if ("$($u.userType)" -eq '') { 'Member' } else { $u.userType })
                AccountEnabled = (Format-Bool $u.accountEnabled)
                OnPremSynced = (Format-Bool $u.onPremisesSyncEnabled)
                UsageLocation = $u.usageLocation; CreatedDateTime = ''
                LicenseCount = $licCount
                IsLicensed = $(if ($licCount -gt 0) { 'TRUE' } else { 'FALSE' })
            })
        }
        $out
    }

    # ---- Roles (flat second sheet of the export) ----------------------------
    $script:SampleAdminUpns = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    Invoke-Dataset 'Roles' {
        $file = Find-Sample 'exportRoleAssignments*.xlsx'
        $sheets = Get-ExcelSheetInfo $file
        $flat = $null
        foreach ($sh in $sheets) {
            $probe = Import-Excel -Path $file -WorksheetName $sh.Name -NoHeader -EndRow 1
            $joined = ($probe[0].PSObject.Properties | ForEach-Object { "$($_.Value)" }) -join '|'
            if ($joined -match 'roleDisplayName') { $flat = $sh.Name; break }
        }
        if ($null -eq $flat) { throw 'flat role assignment sheet not found' }
        $src = Import-Excel -Path $file -WorksheetName $flat
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in $src) {
            if ("$($r.roleDisplayName)" -eq '') { continue }
            [void]$script:SampleAdminUpns.Add([string]$r.userPrincipalName)
            $out.Add([ordered]@{
                RoleName = $r.roleDisplayName; MemberDisplayName = $r.displayName
                MemberUpn = $r.userPrincipalName; MemberType = $r.objectType
                AssignmentType = 'Active'
            })
        }
        $out
    }

    # top-level sample loads are resilient: a missing file must only fail the
    # datasets that need it, never the whole run
    $dlMembers = @()
    try { $dlMembers = Import-Excel (Find-Sample 'DLGroupMember.xlsx') }
    catch { Add-Warning ("Sample: " + $_.Exception.Message) }

    # ---- Groups (synthesized: sample kit has no full group export) ----------
    Invoke-Dataset 'Groups' {
        $out = New-Object System.Collections.Generic.List[object]
        # Microsoft 365 groups: one per team in the Teams admin export
        $list = Import-Excel (Find-Sample 'TeamsList_*.xlsx')
        foreach ($t in $list) {
            $name = Get-Trimmed $t 'Name'
            if ("$name" -eq '') { continue }
            $out.Add([ordered]@{
                Id = [string](Get-Trimmed $t 'Groups Id'); DisplayName = $name; Mail = ''
                Category = 'Microsoft 365'; Visibility = (Get-Trimmed $t 'Privacy')
                IsTeam = 'TRUE'
            })
        }
        # Distribution groups: one per DL in the legacy member export
        $dlNames = $dlMembers | Group-Object DLName
        foreach ($g in $dlNames) {
            $out.Add([ordered]@{
                Id = ''; DisplayName = ($g.Name -split '@')[0]; Mail = $g.Name
                Category = 'Distribution'; Visibility = ''; IsTeam = 'FALSE'
            })
        }
        # Security groups: deterministic synthetic set (no sample export exists)
        for ($i = 1; $i -le 20; $i++) {
            $out.Add([ordered]@{
                Id = ''; DisplayName = ('Sample Security Group {0:d2}' -f $i); Mail = ''
                Category = 'Security'; Visibility = ''; IsTeam = 'FALSE'
            })
        }
        $out
    }

    # ---- MFA (synthesized deterministically; sample kit has no MFA export) --
    Invoke-Dataset 'MFA' {
        $src = Import-Csv (Find-Sample 'exportUsers_*.csv')
        $out = New-Object System.Collections.Generic.List[object]
        $i = 0
        foreach ($u in $src) {
            $i++
            $isAdmin = $script:SampleAdminUpns.Contains([string]$u.userPrincipalName)
            $registered = (($i % 3) -ne 0)   # deterministic: 2 of 3 users registered
            $methods = ''
            if ($registered) { $methods = 'microsoftAuthenticatorPush;softwareOneTimePasscode' }
            $out.Add([ordered]@{
                UserPrincipalName = $u.userPrincipalName
                IsMfaRegistered = (Format-Bool $registered)
                IsMfaCapable = (Format-Bool $registered)
                IsAdmin = (Format-Bool $isAdmin)
                MethodsRegistered = $methods
            })
        }
        $out
    }

    # ---- Enterprise apps -----------------------------------------------------
    Invoke-Dataset 'EnterpriseApps' {
        $src = Import-Excel (Find-Sample 'EnterpriseAppsList.xlsx')
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($a in $src) {
            if ("$($a.displayName)" -eq '') { continue }
            $out.Add([ordered]@{
                DisplayName = $a.displayName; AppId = $a.appId; PublisherName = ''
                CreatedDateTime = (Format-IsoDate $a.createdDateTime)
                AccountEnabled = (Format-Bool $a.accountEnabled)
                Homepage = $a.homepageUrl
                IsMicrosoftFirstParty = (Format-Bool ("$($a.applicationType)" -eq 'Microsoft Application'))
            })
        }
        $out
    }

    # ---- App registrations ----------------------------------------------------
    Invoke-Dataset 'AppRegistrations' {
        $src = Import-Excel (Find-Sample 'AppRegistrationList.xlsx')
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($a in $src) {
            if ("$($a.displayName)" -eq '') { continue }
            $pw = [string]$a.passwordCredentials
            $kc = [string]$a.keyCredentials
            $secretCount = ([regex]::Matches($pw, '"keyId"')).Count
            $certCount = ([regex]::Matches($kc, '"keyId"')).Count
            $nearest = ''
            $dates = [regex]::Matches($pw, '"endDateTime"\s*:\s*"([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
            if ($dates) {
                $parsed = @()
                foreach ($d in $dates) { $p = Format-IsoDate $d; if ($p -ne '') { $parsed += $p } }
                if ($parsed.Count -gt 0) { $nearest = ($parsed | Sort-Object | Select-Object -First 1) }
            }
            $out.Add([ordered]@{
                DisplayName = $a.displayName; AppId = $a.appId
                CreatedDateTime = (Format-IsoDate $a.createdDateTime)
                SecretCount = $secretCount; NearestSecretExpiry = $nearest; CertCount = $certCount
            })
        }
        $out
    }

    # ---- Licenses ----------------------------------------------------------------
    Invoke-Dataset 'Licenses' {
        $src = Import-Excel (Find-Sample 'ProductList_*.xlsx')
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($p in $src) {
            $title = Get-Trimmed $p 'Product Title'
            if ("$title" -eq '') { continue }
            $total = [int](Get-Trimmed $p 'Total Licenses')
            $assigned = [int](Get-Trimmed $p 'Assigned licenses')
            $out.Add([ordered]@{
                SkuPartNumber = (("$title").ToUpper() -replace '[^A-Z0-9]+', '_')
                FriendlyName = $title; Total = $total; Assigned = $assigned
                Available = ($total - $assigned)
            })
        }
        $out
    }

    # ---- Mailboxes ------------------------------------------------------------------
    # In sample mode the run reference date is anchored to the sample data
    # (max LastUserActionTime + 5 days) so the activity bands are exercised.
    $script:SampleMaxAction = [datetime]::MinValue
    $mbxTypes = @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','SchedulingMailbox','GroupMailbox')
    Invoke-Dataset 'Mailboxes' {
        $src = Import-Excel (Find-Sample 'MBXStats_*.xlsx')
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($m in $src) {
            if ($mbxTypes -notcontains "$($m.RecipientTypeDetails)") { continue }
            $isoAction = Convert-LegacyDate $m.LastUserActionTime
            if ($isoAction -ne '') {
                $d = [datetime]::ParseExact($isoAction, 'yyyy-MM-dd', $script:Inv)
                if ($d -gt $script:SampleMaxAction) { $script:SampleMaxAction = $d }
            }
            $size = ''
            if ("$($m.'MBX Size (GB)')" -ne 'N/A' -and "$($m.'MBX Size (GB)')" -ne '') { $size = Format-Num $m.'MBX Size (GB)' 3 }
            $out.Add([ordered]@{
                DisplayName = $m.DisplayName; PrimarySmtpAddress = $m.PrimarySmtpAddress
                RecipientTypeDetails = $m.RecipientTypeDetails; SizeGB = $size; ItemCount = ''
                LastUserActionTime = (Convert-LegacyDate $m.LastUserActionTime)
                LastEmailReceived = (Convert-LegacyDate $m.LastEmailReceived)
                LastEmailSent = (Convert-LegacyDate $m.LastEmailSent)
                ArchiveEnabled = (Format-Bool $m.IsArchiveMailbox)
                LitigationHold = 'FALSE'
                ForwardingSmtpAddress = ''; ForwardingAddress = ''
                ForwardsExternally = 'FALSE'   # legacy export has no forwarding data
            })
        }
        $out
    }

    # ---- Mailbox permissions --------------------------------------------------------
    Invoke-Dataset 'MailboxPermissions' {
        $out = New-Object System.Collections.Generic.List[object]
        $fa = Import-Csv (Find-Sample 'SMB_FullAccess_Permissions.csv') -Delimiter ';'
        foreach ($r in $fa) { $out.Add([ordered]@{ Mailbox = $r.MBX; PermissionType = 'FullAccess'; Grantee = $r.User }) }
        $sa = Import-Csv (Find-Sample 'SMB_SendAS_Perms.csv') -Delimiter ';'
        foreach ($r in $sa) { $out.Add([ordered]@{ Mailbox = $r.MBX; PermissionType = 'SendAs'; Grantee = $r.User }) }
        $sobFile = Get-ChildItem -Path $SamplePath -Recurse -Filter 'SMB_SOB_Perms.csv' | Select-Object -First 1
        if ($sobFile -and $sobFile.Length -gt 0) {
            $sob = Import-Csv $sobFile.FullName -Delimiter ';'
            foreach ($r in $sob) { $out.Add([ordered]@{ Mailbox = $r.MBX; PermissionType = 'SendOnBehalf'; Grantee = $r.User }) }
        }
        $out
    }

    # ---- Distribution groups + members ------------------------------------------------
    Invoke-Dataset 'DistributionGroups' {
        $out = New-Object System.Collections.Generic.List[object]
        $groups = $dlMembers | Group-Object DLName
        foreach ($g in $groups) {
            $out.Add([ordered]@{
                DisplayName = ($g.Name -split '@')[0]; PrimarySmtpAddress = $g.Name
                Type = 'Distribution'; MemberCountDirect = $g.Count
            })
        }
        $out
    }
    Invoke-Dataset 'DLMembers' {
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($m in $dlMembers) {
            $out.Add([ordered]@{
                DLName = $m.DLName; MemberDisplayName = $m.UserDisplayName
                MemberSmtp = $m.UserPrimarySmtpAddress; MemberRecipientType = $m.RecipientType
            })
        }
        $out
    }

    # ---- Domains --------------------------------------------------------------------------
    Invoke-Dataset 'Domains' {
        $src = Import-Csv (Find-Sample 'DomainList_*.csv') -Delimiter ';'
        $out = New-Object System.Collections.Generic.List[object]
        $defaultSet = $false
        foreach ($d in $src) {
            $isDefault = 'FALSE'
            if (-not $defaultSet -and $d.Type -eq 'Authoritative' -and $d.DomainName -notlike '*.onmicrosoft.com') {
                $isDefault = 'TRUE'; $defaultSet = $true
            }
            $out.Add([ordered]@{
                DomainName = $d.DomainName; Type = $d.Type; IsDefault = $isDefault
                DkimEnabled = 'FALSE'; PasswordValidityDays = 2147483647
            })
        }
        $out
    }

    # ---- Transport rules ------------------------------------------------------------------
    Invoke-Dataset 'TransportRules' {
        $src = Import-Csv (Find-Sample 'TransportRules_*.csv') -Delimiter ';'
        $out = New-Object System.Collections.Generic.List[object]
        $i = 0
        foreach ($r in $src) {
            $out.Add([ordered]@{ Name = $r.Name; State = $r.State; Priority = $i; Comments = $r.Desc })
            $i++
        }
        $out
    }

    # ---- Email activity (D30 counts) ---------------------------------------------------------
    Invoke-Dataset 'EmailActivity' {
        $src = Import-Excel (Find-Sample 'EmailActivityCounts*.xlsx')
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in $src) {
            $rd = Get-Trimmed $r 'Report Date'
            if ($null -eq $rd) { continue }
            $out.Add([ordered]@{
                ReportDate = (Format-IsoDate $rd)
                Send = (Format-Num (Get-Trimmed $r 'Send'))
                Receive = (Format-Num (Get-Trimmed $r 'Receive'))
                Read = (Format-Num (Get-Trimmed $r 'Read'))
            })
        }
        $out
    }

    # ---- Email user activity (synthesized from mailbox stats) -----------------------------------
    Invoke-Dataset 'EmailUserActivity' {
        $src = Import-Excel (Find-Sample 'MBXStats_*.xlsx')
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($m in $src) {
            if ("$($m.RecipientTypeDetails)" -ne 'UserMailbox') { continue }
            $sendCount = 0; $recvCount = 0
            if ("$($m.LastEmailSent)" -notlike 'No*' -and "$($m.LastEmailSent)" -ne '') { $sendCount = 10 }
            if ("$($m.LastEmailReceived)" -notlike 'No*' -and "$($m.LastEmailReceived)" -ne '') { $recvCount = 25 }
            $out.Add([ordered]@{
                UserPrincipalName = $m.PrimarySmtpAddress
                LastActivityDate = (Convert-LegacyDate $m.LastUserActionTime)
                SendCount = $sendCount; ReceiveCount = $recvCount; ReadCount = ($recvCount * 2)
            })
        }
        $out
    }

    # ---- SharePoint sites --------------------------------------------------------------------------
    Invoke-Dataset 'SPOSites' {
        $src = Import-Excel (Find-Sample 'SPStats.xlsx')
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($s in $src) {
            $name = Get-Trimmed $s 'Site name'
            if ($null -eq $name) {
                # header may carry a BOM: first property is the site name
                $name = $s.PSObject.Properties.Value | Select-Object -First 1
            }
            $tmpl = [string](Get-Trimmed $s 'Template')
            $out.Add([ordered]@{
                SiteUrl = (Get-Trimmed $s 'URL'); OwnerDisplayName = (Get-Trimmed $s 'Created by')
                LastActivityDate = (Format-IsoDate (Get-Trimmed $s 'Last activity (UTC)'))
                FileCount = (Format-Num (Get-Trimmed $s 'Files'))
                ActiveFileCount = 0
                StorageUsedGB = (Format-Num (Get-Trimmed $s 'Storage used (GB)') 3)
                RootWebTemplate = $tmpl
                IsTeamsConnected = (Format-Bool ($tmpl -eq 'Team site'))
            })
        }
        $out
    }

    # ---- OneDrive ------------------------------------------------------------------------------------
    Invoke-Dataset 'OneDrive' {
        $src = Import-Excel (Find-Sample 'OneDriveUsageAccountDetail*.xlsx')
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($o in $src) {
            $upn = Get-Trimmed $o 'Owner Principal Name'
            if ("$upn" -eq '') { continue }
            $bytes = [double](Get-Trimmed $o 'Storage Used (Byte)')
            $out.Add([ordered]@{
                OwnerUpn = $upn; OwnerDisplayName = (Get-Trimmed $o 'Owner Display Name')
                LastActivityDate = (Format-IsoDate (Get-Trimmed $o 'Last Activity Date'))
                FileCount = (Format-Num (Get-Trimmed $o 'File Count'))
                ActiveFileCount = (Format-Num (Get-Trimmed $o 'Active File Count'))
                StorageUsedGB = (Format-Num ($bytes / 1GB) 3)
            })
        }
        $out
    }

    # ---- Teams -----------------------------------------------------------------------------------------
    Invoke-Dataset 'Teams' {
        $list = Import-Excel (Find-Sample 'TeamsList_*.xlsx')
        $act = Import-Excel (Find-Sample 'TeamsTeamActivityDetail*.xlsx')
        $actById = @{}
        foreach ($a in $act) {
            $tid = [string](Get-Trimmed $a 'Team Id')
            if ($tid -ne '') { $actById[$tid] = $a }
        }
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($t in $list) {
            $name = Get-Trimmed $t 'Name'
            if ("$name" -eq '') { continue }
            $gid = [string](Get-Trimmed $t 'Groups Id')
            $active = ''; $msgs = ''; $lastAct = ''
            if ($actById.ContainsKey($gid)) {
                $a = $actById[$gid]
                $active = Format-Num (Get-Trimmed $a 'Active Users')
                $msgs = Format-Num (Get-Trimmed $a 'Channel Messages')
                $lastAct = Format-IsoDate (Get-Trimmed $a 'Last Activity Date')
            }
            $out.Add([ordered]@{
                TeamName = $name; GroupId = $gid; Visibility = (Get-Trimmed $t 'Privacy')
                MemberCount = (Format-Num (Get-Trimmed $t 'Team Members'))
                ActiveUsers90d = $active; ChannelMessages90d = $msgs; LastActivityDate = $lastAct
            })
        }
        $out
    }

    # ---- Teams user activity (sorted desc so the chart shows the top-30 users) --------------------------
    Invoke-Dataset 'TeamsUserActivity' {
        $src = Import-Excel (Find-Sample 'TeamsUserActivityUserDetail*.xlsx')
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($u in $src) {
            $upn = Get-Trimmed $u 'User Principal Name'
            if ("$upn" -eq '') { continue }
            $tc = [int](Get-Trimmed $u 'Team Chat Message Count')
            $pc = [int](Get-Trimmed $u 'Private Chat Message Count')
            $ca = [int](Get-Trimmed $u 'Call Count')
            $me = [int](Get-Trimmed $u 'Meeting Count')
            $rows.Add([pscustomobject]@{
                UserPrincipalName = $upn
                LastActivityDate = (Format-IsoDate (Get-Trimmed $u 'Last Activity Date'))
                TeamChatMessages = $tc; PrivateChatMessages = $pc; Calls = $ca; Meetings = $me
                Total = ($tc + $pc + $ca + $me)
            })
        }
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in ($rows | Sort-Object Total, UserPrincipalName -Descending)) {
            $out.Add([ordered]@{
                UserPrincipalName = $r.UserPrincipalName; LastActivityDate = $r.LastActivityDate
                TeamChatMessages = $r.TeamChatMessages; PrivateChatMessages = $r.PrivateChatMessages
                Calls = $r.Calls; Meetings = $r.Meetings
            })
        }
        $out
    }

    # ---- Devices ------------------------------------------------------------------------------------------
    $intune = @()
    try { $intune = Import-Excel (Find-Sample 'DevicesWithInventory_*.xlsx') }
    catch { Add-Warning ("Sample: " + $_.Exception.Message) }
    Invoke-Dataset 'DevicesEntra' {
        $out = New-Object System.Collections.Generic.List[object]
        $i = 0
        foreach ($d in $intune) {
            if ("$($d.'Device name')" -eq '') { continue }
            $i++
            $managedBy = [string]$d.'Managed by'
            $out.Add([ordered]@{
                DisplayName = $d.'Device name'; OS = 'Windows'; OSVersion = $d.'OS version'
                TrustType = 'AzureAd'
                LastSignIn = (Format-IsoDate $d.'Last check-in')
                IsCompliant = (Format-Bool (($i % 7) -ne 0))
                IsManaged = (Format-Bool ($managedBy -match 'Intune|MDM|ConfigMgr'))
                RegisteredDateTime = (Format-IsoDate $d.'Enrollment date')
            })
        }
        $out
    }
    Invoke-Dataset 'DevicesIntune' {
        $out = New-Object System.Collections.Generic.List[object]
        $i = 0
        foreach ($d in $intune) {
            if ("$($d.'Device name')" -eq '') { continue }
            $i++
            $compliance = 'compliant'
            if (($i % 10) -eq 0) { $compliance = 'noncompliant' }
            $out.Add([ordered]@{
                DeviceName = $d.'Device name'; OS = 'Windows'; OSVersion = $d.'OS version'
                ComplianceState = $compliance
                LastSyncDateTime = (Format-IsoDate $d.'Last check-in')
                ManagementAgent = $d.'Managed by'; Manufacturer = $d.Manufacturer; Model = $d.Model
                EnrolledDateTime = (Format-IsoDate $d.'Enrollment date')
            })
        }
        $out
    }

    # ---- Security settings (synthesized single row) -----------------------------------------------------------
    Invoke-Dataset 'SecuritySettings' {
        ,([ordered]@{
            SecurityDefaultsEnabled = 'FALSE'; CAPoliciesTotal = 0; CAPoliciesEnabled = 0
            TenantSharingCapability = 'ExternalUserAndGuestSharing'
            SmtpAuthDisabled = 'FALSE'; AuditEnabled = 'TRUE'
        })
    }

    # ---- MANUAL landing files (exercise the M_ sheets import path) --------------------------------------------
    Write-Log "Writing MANUAL sample files ..."
    try {
        $teamsCols = @('Name','Standard Channels','Private Channels','Shared Channels','Team Members','Owners','Guests','Privacy','Status','Classification','Groups Id','Expiration Date','Description','Sensitivity Label')
        $list = Import-Excel (Find-Sample 'TeamsList_*.xlsx')
        $manualTeams = New-Object System.Collections.Generic.List[object]
        foreach ($t in $list) {
            if ("$(Get-Trimmed $t 'Name')" -eq '') { continue }
            $row = [ordered]@{}
            foreach ($c in $teamsCols) { $row[$c] = Get-Trimmed $t $c }
            $manualTeams.Add($row)
        }
        Write-AuditCsv -FileName 'TeamsAdminExport.csv' -Columns $teamsCols -Rows $manualTeams -Folder $script:ManualPath
    } catch { Add-Warning ("Sample MANUAL TeamsAdminExport: " + $_.Exception.Message) }

    try {
        $spCols = @('Site name','URL','Storage used (GB)','Hub','Template','Last activity (UTC)','Created by','Files','External sharing')
        $sp = Import-Excel (Find-Sample 'SPStats.xlsx')
        $manualSp = New-Object System.Collections.Generic.List[object]
        foreach ($s in $sp) {
            $row = [ordered]@{}
            foreach ($c in $spCols) { $row[$c] = Get-Trimmed $s $c }
            if ("$($row['Site name'])" -eq '') { $row['Site name'] = ($s.PSObject.Properties.Value | Select-Object -First 1) }
            $manualSp.Add($row)
        }
        Write-AuditCsv -FileName 'SPAdminSites.csv' -Columns $spCols -Rows $manualSp -Folder $script:ManualPath
    } catch { Add-Warning ("Sample MANUAL SPAdminSites: " + $_.Exception.Message) }

    $mdeCols = @('Device ID','Device Name','Device Category','Device Type','Device Subtype','Discovery sources','Domain','AAD Device Id','First Seen','Last device update','OS Platform','OS Distribution','OS Version','OS Build','Windows 10 Version','Tags','Group','Is AAD Joined','Device IPs','Device MACs','Risk Level','Exposure Level','Health Status','Onboarding Status','Device Role','Cloud Platforms','Is Internet Facing','Enrollment Status Code','Managed By','Enrollment Status','Vendor','Model')
    $manualMde = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($d in $intune) {
        if ("$($d.'Device name')" -eq '') { continue }
        $i++
        $row = [ordered]@{}
        foreach ($c in $mdeCols) { $row[$c] = '' }
        $row['Device ID'] = 'sample{0:d4}' -f $i
        $row['Device Name'] = [string]$d.'Device name'
        $row['OS Platform'] = 'Windows11'
        $row['Onboarding Status'] = 'Onboarded'
        $row['Managed By'] = [string]$d.'Managed by'
        $manualMde.Add($row)
        if (($i % 8) -eq 0) {
            $disc = [ordered]@{}
            foreach ($c in $mdeCols) { $disc[$c] = '' }
            $disc['Device ID'] = 'sampled{0:d4}' -f $i
            $disc['Device Name'] = ([string]$d.'Device name') + '-iot'
            $disc['OS Platform'] = 'Other'
            $disc['Onboarding Status'] = 'Can be onboarded'
            $manualMde.Add($disc)
        }
    }
    Write-AuditCsv -FileName 'MDE_Devices.csv' -Columns $mdeCols -Rows $manualMde -Folder $script:ManualPath

    $script:TenantId = '00000000-0000-0000-0000-000000000000'
    $script:TenantName = 'Sample Tenant (offline)'
    $script:Account = 'sample@offline.local'
    if ($script:SampleMaxAction -gt [datetime]::MinValue) {
        # anchor the run date to the sample data so activity bands have content
        $script:RunStart = $script:SampleMaxAction.AddDays(5)
        Write-Log ("Sample mode: run reference date anchored to {0:yyyy-MM-dd}" -f $script:RunStart)
    }
    Add-Warning 'Sample data mode: MFA, security settings, device compliance and email user activity are synthesized.'
    Write-RunMeta 'OK'
}

if ($SampleData -ne '') {
    try {
        Convert-SampleData -SamplePath $SampleData
        exit 0
    } catch {
        Write-Log ("SAMPLE MODE FAILED: " + $_.Exception.Message) 'ERROR'
        Write-RunMeta 'FAILED'
        exit 1
    }
}

# ===========================================================================
# LIVE MODE
# ===========================================================================

# ---------------------------------------------------------------------------
# Modules
# ---------------------------------------------------------------------------
function Confirm-Module {
    param([string]$Name, [version]$MinVersion = '0.0')
    $mod = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $mod -or $mod.Version -lt $MinVersion) {
        Write-Log "Installing module $Name (CurrentUser) ..."
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber -MinimumVersion $MinVersion
    }
}

$GraphScopes = @(
    'User.Read.All','Group.Read.All','GroupMember.Read.All','Directory.Read.All',
    'Organization.Read.All','Application.Read.All','Policy.Read.All','Reports.Read.All',
    'ReportSettings.Read.All','AuditLog.Read.All','UserAuthenticationMethod.Read.All',
    'Device.Read.All','DeviceManagementManagedDevices.Read.All','RoleManagement.Read.Directory',
    'Domain.Read.All','SharePointTenantSettings.Read.All','Team.ReadBasic.All'
)

try {
    if (-not $SkipGraph) {
        # only the Authentication sub-module is needed: everything uses Invoke-MgGraphRequest
        Confirm-Module -Name 'Microsoft.Graph.Authentication'
        Import-Module Microsoft.Graph.Authentication
    }
    if (-not $SkipExchange) {
        Confirm-Module -Name 'ExchangeOnlineManagement' -MinVersion '3.4.0'
        Import-Module ExchangeOnlineManagement
    }
} catch {
    Write-Log ("Module installation failed: " + $_.Exception.Message) 'ERROR'
    Write-Log "Fix: open PowerShell and run: Install-Module Microsoft.Graph.Authentication, ExchangeOnlineManagement -Scope CurrentUser" 'ERROR'
    Write-RunMeta 'FAILED'
    exit 1
}

# ---------------------------------------------------------------------------
# Graph helpers
# ---------------------------------------------------------------------------
$script:ScopeHints = @{
    '/users'                 = 'User.Read.All'
    '/groups'                = 'Group.Read.All + GroupMember.Read.All'
    '/reports/'              = 'Reports.Read.All'
    '/admin/reportSettings'  = 'ReportSettings.Read.All'
    '/directoryRoles'        = 'Directory.Read.All'
    '/roleManagement'        = 'RoleManagement.Read.Directory'
    '/servicePrincipals'     = 'Application.Read.All'
    '/applications'          = 'Application.Read.All'
    '/subscribedSkus'        = 'Organization.Read.All'
    '/domains'               = 'Domain.Read.All'
    '/devices'               = 'Device.Read.All'
    '/deviceManagement'      = 'DeviceManagementManagedDevices.Read.All'
    '/policies/'             = 'Policy.Read.All'
    '/identity/conditionalAccess' = 'Policy.Read.All'
    '/admin/sharepoint'      = 'SharePointTenantSettings.Read.All'
}

function Get-ScopeHint {
    param([string]$Uri)
    foreach ($k in $script:ScopeHints.Keys) {
        if ($Uri -like "*$k*") { return $script:ScopeHints[$k] }
    }
    return 'see the scope list in README.md'
}

function Invoke-GraphGet {
    # Paged GET returning all items of .value (or the raw object when no value array)
    param([string]$Uri, [hashtable]$Headers = $null)
    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $attempt = 0
        while ($true) {
            try {
                if ($Headers) { $resp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $Headers -OutputType PSObject }
                else { $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject }
                break
            } catch {
                $msg = $_.Exception.Message
                if ($msg -match '(?i)(throttl|429|too many requests)' -and $attempt -lt 5) {
                    $attempt++
                    $delay = [math]::Min(60, 5 * [math]::Pow(2, $attempt))
                    Write-Log "Throttled on $next, retrying in ${delay}s (attempt $attempt)" 'WARN'
                    Start-Sleep -Seconds $delay
                    continue
                }
                if ($msg -match '(?i)(authorization|access.?denied|forbidden|403|insufficient)') {
                    throw ("Access denied on {0}. Missing delegated consent for: {1}. A Global Admin of the target tenant must consent once (see README, section 'One-time consent')." -f $next, (Get-ScopeHint $next))
                }
                throw
            }
        }
        $hasValue = $false
        if ($resp -is [psobject] -and $resp.PSObject.Properties['value']) { $hasValue = $true }
        if ($hasValue) { foreach ($it in $resp.value) { $items.Add($it) } }
        else { $items.Add($resp) }
        $next = $null
        if ($resp -is [psobject] -and $resp.PSObject.Properties['@odata.nextLink']) {
            $next = $resp.'@odata.nextLink'
        }
    }
    return $items
}

function Get-GraphReportCsv {
    # Downloads a usage report (CSV endpoint) and returns parsed rows
    param([string]$ReportCall)
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("m365rep_" + [guid]::NewGuid().ToString('N') + ".csv")
    try {
        Invoke-MgGraphRequest -Method GET -Uri ("v1.0/reports/{0}?`$format=text/csv" -f $ReportCall) -OutputFilePath $tmp
        return @(Import-Csv $tmp)
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '(?i)(authorization|access.?denied|forbidden|403)') {
            throw ("Access denied on reports/{0}. Missing delegated consent for Reports.Read.All (admin consent needed once)." -f $ReportCall)
        }
        throw
    } finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

# ---------------------------------------------------------------------------
# Sign-in + preflight
# ---------------------------------------------------------------------------
$script:GraphOk = $false; $script:ExoOk = $false

if (-not $SkipGraph) {
    Write-Log "Connecting to Microsoft Graph (browser sign-in, delegated, read-only scopes) ..."
    try {
        Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        $script:Account = $ctx.Account
        $script:TenantId = $ctx.TenantId
        $script:GraphOk = $true
        Write-Log "Graph connected as $($ctx.Account) on tenant $($ctx.TenantId)"

        $granted = @($ctx.Scopes)
        $missing = @()
        foreach ($s in $GraphScopes) { if ($granted -notcontains $s) { $missing += $s } }
        if ($missing.Count -gt 0) {
            Add-Warning ("Missing Graph scopes (admin consent needed once for 'Microsoft Graph Command Line Tools'): " + ($missing -join ', '))
        }
    } catch {
        $m = $_.Exception.Message
        Write-Log ("Graph sign-in failed: " + $m) 'ERROR'
        if ($m -match 'AADSTS65001|consent') {
            Write-Log "The target tenant has not consented to 'Microsoft Graph Command Line Tools'. A Global Admin must run this script once and tick 'Consent on behalf of your organization', or grant admin consent in Entra portal. See README." 'ERROR'
        }
        Add-Warning ("Graph connection failed: " + $m)
    }

    if ($script:GraphOk) {
        # organization name
        try {
            $org = Invoke-GraphGet 'v1.0/organization?$select=id,displayName'
            if ($org.Count -gt 0) { $script:TenantName = $org[0].displayName }
        } catch { Add-Warning ("Preflight organization: " + $_.Exception.Message) }

        # directory roles of the signed-in account
        try {
            $roles = Invoke-GraphGet 'v1.0/me/memberOf/microsoft.graph.directoryRole?$select=displayName'
            $script:AccountRoles = @($roles | ForEach-Object { $_.displayName })
            Write-Log ("Account directory roles: " + ($script:AccountRoles -join ', '))
            if ($script:AccountRoles -notcontains 'Global Reader') {
                Add-Warning "The signed-in account does not hold Global Reader. Several datasets may fail. Ask the target tenant to assign Global Reader to this account."
            }
        } catch { Add-Warning ("Preflight roles check: " + $_.Exception.Message) }

        # concealed names in usage reports
        try {
            $rs = Invoke-GraphGet 'v1.0/admin/reportSettings'
            if ($rs.Count -gt 0 -and $rs[0].displayConcealedNames) {
                $script:ConcealedNames = $true
                Add-Warning "Usage reports hide user/site names (pseudonymized). Per-user joins will degrade. Fix (target tenant Global Admin): M365 admin center > Settings > Org settings > Reports > untick 'Display concealed user, group, and site names in all reports'."
            }
        } catch { Add-Warning ("Preflight reportSettings: " + $_.Exception.Message) }
    }
}

if (-not $SkipExchange) {
    Write-Log "Connecting to Exchange Online (browser sign-in) ..."
    try {
        if ($UserPrincipalName -ne '') {
            Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop
        } elseif ($script:Account -ne '') {
            Connect-ExchangeOnline -UserPrincipalName $script:Account -ShowBanner:$false -ErrorAction Stop
        } else {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
        $script:ExoOk = $true
        Write-Log "Exchange Online connected."
    } catch {
        Write-Log ("Exchange Online sign-in failed: " + $_.Exception.Message) 'ERROR'
        Add-Warning ("Exchange connection failed: " + $_.Exception.Message)
    }
}

if (-not $script:GraphOk -and -not $script:ExoOk) {
    Write-Log "No connection could be established. Stopping." 'ERROR'
    Write-RunMeta 'FAILED'
    exit 1
}

# ---------------------------------------------------------------------------
# GRAPH DATASETS
# ---------------------------------------------------------------------------
$script:RawGroups = @()
$script:EmailUserMap = @{}

if ($script:GraphOk) {

    Invoke-Dataset 'Users' {
        $users = Invoke-GraphGet 'v1.0/users?$select=id,displayName,userPrincipalName,mail,userType,accountEnabled,onPremisesSyncEnabled,usageLocation,createdDateTime,assignedLicenses&$top=999'
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($u in $users) {
            $lic = 0
            if ($u.assignedLicenses) { $lic = @($u.assignedLicenses).Count }
            $ut = [string]$u.userType
            if ($ut -eq '') { $ut = 'Member' }
            $out.Add([ordered]@{
                Id = $u.id; DisplayName = $u.displayName; UserPrincipalName = $u.userPrincipalName
                Mail = $u.mail; UserType = $ut
                AccountEnabled = (Format-Bool $u.accountEnabled)
                OnPremSynced = (Format-Bool $u.onPremisesSyncEnabled)
                UsageLocation = $u.usageLocation
                CreatedDateTime = (Format-IsoDate $u.createdDateTime)
                LicenseCount = $lic
                IsLicensed = $(if ($lic -gt 0) { 'TRUE' } else { 'FALSE' })
            })
        }
        $out
    }

    Invoke-Dataset 'Groups' {
        $script:RawGroups = Invoke-GraphGet 'v1.0/groups?$select=id,displayName,mail,mailEnabled,securityEnabled,groupTypes,visibility,resourceProvisioningOptions&$top=999'
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($g in $script:RawGroups) {
            $isUnified = ($g.groupTypes -contains 'Unified')
            $category = 'Distribution'
            if ($isUnified) { $category = 'Microsoft 365' }
            elseif ($g.securityEnabled -and $g.mailEnabled) { $category = 'Mail-enabled security' }
            elseif ($g.securityEnabled) { $category = 'Security' }
            $isTeam = ($g.resourceProvisioningOptions -contains 'Team')
            $out.Add([ordered]@{
                Id = $g.id; DisplayName = $g.displayName; Mail = $g.mail
                Category = $category; Visibility = $g.visibility
                IsTeam = (Format-Bool $isTeam)
            })
        }
        $out
    }

    Invoke-Dataset 'MFA' {
        # no $top: this endpoint does not reliably accept large page sizes;
        # Invoke-GraphGet follows @odata.nextLink anyway
        $regs = Invoke-GraphGet 'v1.0/reports/authenticationMethods/userRegistrationDetails'
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in $regs) {
            $methods = ''
            if ($r.methodsRegistered) { $methods = (@($r.methodsRegistered) -join ';') }
            $out.Add([ordered]@{
                UserPrincipalName = $r.userPrincipalName
                IsMfaRegistered = (Format-Bool $r.isMfaRegistered)
                IsMfaCapable = (Format-Bool $r.isMfaCapable)
                IsAdmin = (Format-Bool $r.isAdmin)
                MethodsRegistered = $methods
            })
        }
        $out
    }

    Invoke-Dataset 'Roles' {
        $out = New-Object System.Collections.Generic.List[object]
        $activeRoles = Invoke-GraphGet 'v1.0/directoryRoles?$select=id,displayName'
        foreach ($role in $activeRoles) {
            $members = @()
            try {
                $members = Invoke-GraphGet ("v1.0/directoryRoles/{0}/members?`$select=id,displayName,userPrincipalName&`$top=999" -f $role.id)
            } catch {
                Add-Warning ("Roles: cannot list members of '{0}': {1}" -f $role.displayName, $_.Exception.Message)
                continue
            }
            foreach ($mm in $members) {
                $otype = 'unknown'
                if ($mm.'@odata.type') { $otype = ([string]$mm.'@odata.type').Split('.')[-1] }
                $upn = ''
                if ($mm.PSObject.Properties['userPrincipalName']) { $upn = $mm.userPrincipalName }
                $out.Add([ordered]@{
                    RoleName = $role.displayName; MemberDisplayName = $mm.displayName
                    MemberUpn = $upn; MemberType = $otype; AssignmentType = 'Active'
                })
            }
        }
        # PIM-eligible assignments (requires Entra P2; optional)
        try {
            $defs = Invoke-GraphGet 'v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName'
            $defMap = @{}
            foreach ($d in $defs) { $defMap[[string]$d.id] = $d.displayName }
            $eligible = Invoke-GraphGet 'v1.0/roleManagement/directory/roleEligibilitySchedules?$expand=principal'
            foreach ($e in $eligible) {
                $rn = $defMap[[string]$e.roleDefinitionId]
                if ("$rn" -eq '') { $rn = $e.roleDefinitionId }
                $pdisp = ''; $pupn = ''; $ptype = 'unknown'
                if ($e.principal) {
                    $pdisp = $e.principal.displayName
                    if ($e.principal.PSObject.Properties['userPrincipalName']) { $pupn = $e.principal.userPrincipalName }
                    if ($e.principal.'@odata.type') { $ptype = ([string]$e.principal.'@odata.type').Split('.')[-1] }
                }
                $out.Add([ordered]@{
                    RoleName = $rn; MemberDisplayName = $pdisp; MemberUpn = $pupn
                    MemberType = $ptype; AssignmentType = 'Eligible'
                })
            }
        } catch {
            Add-Warning ("Roles: PIM eligible assignments not readable (needs Entra P2 / PIM): " + $_.Exception.Message)
        }
        $out
    }

    Invoke-Dataset 'EnterpriseApps' {
        $msTenants = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')
        $sps = Invoke-GraphGet 'v1.0/servicePrincipals?$filter=servicePrincipalType eq ''Application''&$select=displayName,appId,publisherName,createdDateTime,accountEnabled,homepage,appOwnerOrganizationId&$count=true&$top=999' @{ ConsistencyLevel = 'eventual' }
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($sp in $sps) {
            $isMs = ($msTenants -contains [string]$sp.appOwnerOrganizationId)
            $out.Add([ordered]@{
                DisplayName = $sp.displayName; AppId = $sp.appId; PublisherName = $sp.publisherName
                CreatedDateTime = (Format-IsoDate $sp.createdDateTime)
                AccountEnabled = (Format-Bool $sp.accountEnabled)
                Homepage = $sp.homepage
                IsMicrosoftFirstParty = (Format-Bool $isMs)
            })
        }
        $out
    }

    Invoke-Dataset 'AppRegistrations' {
        $apps = Invoke-GraphGet 'v1.0/applications?$select=displayName,appId,createdDateTime,passwordCredentials,keyCredentials&$top=999'
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($a in $apps) {
            $pw = @(); if ($a.passwordCredentials) { $pw = @($a.passwordCredentials) }
            $kc = @(); if ($a.keyCredentials) { $kc = @($a.keyCredentials) }
            $nearest = ''
            if ($pw.Count -gt 0) {
                $dates = @()
                foreach ($cred in $pw) { $d = Format-IsoDate $cred.endDateTime; if ($d -ne '') { $dates += $d } }
                if ($dates.Count -gt 0) { $nearest = ($dates | Sort-Object | Select-Object -First 1) }
            }
            $out.Add([ordered]@{
                DisplayName = $a.displayName; AppId = $a.appId
                CreatedDateTime = (Format-IsoDate $a.createdDateTime)
                SecretCount = $pw.Count; NearestSecretExpiry = $nearest; CertCount = $kc.Count
            })
        }
        $out
    }

    Invoke-Dataset 'Licenses' {
        $skus = Invoke-GraphGet 'v1.0/subscribedSkus'
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($sku in $skus) {
            $total = 0
            if ($sku.prepaidUnits) { $total = [int]$sku.prepaidUnits.enabled }
            $assigned = [int]$sku.consumedUnits
            $out.Add([ordered]@{
                SkuPartNumber = $sku.skuPartNumber
                FriendlyName = (Get-SkuFriendlyName $sku.skuPartNumber)
                Total = $total; Assigned = $assigned; Available = ($total - $assigned)
            })
        }
        $out
    }

    Invoke-Dataset 'EmailActivity' {
        $rows = Get-GraphReportCsv "getEmailActivityCounts(period='D30')"
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rows) {
            $out.Add([ordered]@{
                ReportDate = (Format-IsoDate $r.'Report Date')
                Send = (Format-Num $r.Send); Receive = (Format-Num $r.Receive); Read = (Format-Num $r.Read)
            })
        }
        $out
    }

    Invoke-Dataset 'EmailUserActivity' {
        $rows = Get-GraphReportCsv "getEmailActivityUserDetail(period='D90')"
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rows) {
            $upn = [string]$r.'User Principal Name'
            $item = [ordered]@{
                UserPrincipalName = $upn
                LastActivityDate = (Format-IsoDate $r.'Last Activity Date')
                SendCount = (Format-Num $r.'Send Count')
                ReceiveCount = (Format-Num $r.'Receive Count')
                ReadCount = (Format-Num $r.'Read Count')
            }
            $out.Add($item)
            if ($upn -ne '') { $script:EmailUserMap[$upn.ToLower()] = $item }
        }
        $out
    }

    Invoke-Dataset 'SPOSites' {
        $rows = Get-GraphReportCsv "getSharePointSiteUsageDetail(period='D90')"
        $out = New-Object System.Collections.Generic.List[object]
        $blankUrl = 0
        foreach ($r in $rows) {
            $url = [string]$r.'Site URL'
            if ($url -eq '') { $url = [string]$r.'Site Id'; $blankUrl++ }
            $tmpl = [string]$r.'Root Web Template'
            $bytes = 0.0
            [void][double]::TryParse([string]$r.'Storage Used (Byte)', [System.Globalization.NumberStyles]::Any, $script:Inv, [ref]$bytes)
            $out.Add([ordered]@{
                SiteUrl = $url; OwnerDisplayName = $r.'Owner Display Name'
                LastActivityDate = (Format-IsoDate $r.'Last Activity Date')
                FileCount = (Format-Num $r.'File Count')
                ActiveFileCount = (Format-Num $r.'Active File Count')
                StorageUsedGB = (Format-Num ($bytes / 1GB) 3)
                RootWebTemplate = $tmpl
                IsTeamsConnected = (Format-Bool ($tmpl -eq 'Group'))
            })
        }
        if ($blankUrl -gt 0) {
            Add-Warning "SPOSites: $blankUrl sites have no URL in the usage report (site id used instead). This is normal on recent tenants."
        }
        $out
    }

    Invoke-Dataset 'OneDrive' {
        $rows = Get-GraphReportCsv "getOneDriveUsageAccountDetail(period='D90')"
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rows) {
            if ("$($r.'Is Deleted')" -match '(?i)true') { continue }
            $bytes = 0.0
            [void][double]::TryParse([string]$r.'Storage Used (Byte)', [System.Globalization.NumberStyles]::Any, $script:Inv, [ref]$bytes)
            $out.Add([ordered]@{
                OwnerUpn = $r.'Owner Principal Name'; OwnerDisplayName = $r.'Owner Display Name'
                LastActivityDate = (Format-IsoDate $r.'Last Activity Date')
                FileCount = (Format-Num $r.'File Count')
                ActiveFileCount = (Format-Num $r.'Active File Count')
                StorageUsedGB = (Format-Num ($bytes / 1GB) 3)
            })
        }
        $out
    }

    Invoke-Dataset 'Teams' {
        $teams = @($script:RawGroups | Where-Object { $_.resourceProvisioningOptions -contains 'Team' })
        $actRows = @()
        try { $actRows = Get-GraphReportCsv "getTeamsTeamActivityDetail(period='D90')" }
        catch { Add-Warning ("Teams: activity report failed: " + $_.Exception.Message) }
        $actById = @{}
        foreach ($a in $actRows) {
            $tid = [string]$a.'Team Id'
            if ($tid -ne '') { $actById[$tid] = $a }
        }
        $out = New-Object System.Collections.Generic.List[object]
        $i = 0
        foreach ($t in $teams) {
            $i++
            Write-Progress -Activity 'Teams member counts' -Status "$i of $($teams.Count)" -PercentComplete ([int](100 * $i / [math]::Max(1, $teams.Count)))
            $memberCount = ''
            try {
                $memberCount = Invoke-MgGraphRequest -Method GET -Uri ("v1.0/groups/{0}/members/`$count" -f $t.id) -Headers @{ ConsistencyLevel = 'eventual' } -OutputType HttpResponseMessage
                $memberCount = $memberCount.Content.ReadAsStringAsync().Result.Trim('"')
            } catch { $memberCount = '' }
            $active = ''; $msgs = ''; $lastAct = ''
            if ($actById.ContainsKey([string]$t.id)) {
                $a = $actById[[string]$t.id]
                $active = Format-Num $a.'Active Users'
                $msgs = Format-Num $a.'Channel Messages'
                $lastAct = Format-IsoDate $a.'Last Activity Date'
            }
            $out.Add([ordered]@{
                TeamName = $t.displayName; GroupId = $t.id; Visibility = $t.visibility
                MemberCount = $memberCount; ActiveUsers90d = $active
                ChannelMessages90d = $msgs; LastActivityDate = $lastAct
            })
        }
        Write-Progress -Activity 'Teams member counts' -Completed
        $out
    }

    Invoke-Dataset 'TeamsUserActivity' {
        $rows = Get-GraphReportCsv "getTeamsUserActivityUserDetail(period='D90')"
        $tmp = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rows) {
            $tc = 0; $pc = 0; $ca = 0; $me = 0
            [void][int]::TryParse([string]$r.'Team Chat Message Count', [ref]$tc)
            [void][int]::TryParse([string]$r.'Private Chat Message Count', [ref]$pc)
            [void][int]::TryParse([string]$r.'Call Count', [ref]$ca)
            [void][int]::TryParse([string]$r.'Meeting Count', [ref]$me)
            $tmp.Add([pscustomobject]@{
                UserPrincipalName = $r.'User Principal Name'
                LastActivityDate = (Format-IsoDate $r.'Last Activity Date')
                TeamChatMessages = $tc; PrivateChatMessages = $pc; Calls = $ca; Meetings = $me
                Total = ($tc + $pc + $ca + $me)
            })
        }
        # sorted desc so the dashboard bar chart (rows 2-31) shows the 30 most active users
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in ($tmp | Sort-Object Total -Descending)) {
            $out.Add([ordered]@{
                UserPrincipalName = $r.UserPrincipalName; LastActivityDate = $r.LastActivityDate
                TeamChatMessages = $r.TeamChatMessages; PrivateChatMessages = $r.PrivateChatMessages
                Calls = $r.Calls; Meetings = $r.Meetings
            })
        }
        $out
    }

    Invoke-Dataset 'DevicesEntra' {
        $devs = Invoke-GraphGet 'v1.0/devices?$select=displayName,operatingSystem,operatingSystemVersion,trustType,approximateLastSignInDateTime,isCompliant,isManaged,registrationDateTime&$top=999'
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($d in $devs) {
            $out.Add([ordered]@{
                DisplayName = $d.displayName; OS = $d.operatingSystem; OSVersion = $d.operatingSystemVersion
                TrustType = $d.trustType
                LastSignIn = (Format-IsoDate $d.approximateLastSignInDateTime)
                IsCompliant = (Format-Bool $d.isCompliant)
                IsManaged = (Format-Bool $d.isManaged)
                RegisteredDateTime = (Format-IsoDate $d.registrationDateTime)
            })
        }
        $out
    }

    Invoke-Dataset 'DevicesIntune' {
        # no $top: Intune endpoints page at their own size; nextLink is followed
        $devs = Invoke-GraphGet 'v1.0/deviceManagement/managedDevices?$select=deviceName,operatingSystem,osVersion,complianceState,lastSyncDateTime,managementAgent,manufacturer,model,enrolledDateTime'
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($d in $devs) {
            $out.Add([ordered]@{
                DeviceName = $d.deviceName; OS = $d.operatingSystem; OSVersion = $d.osVersion
                ComplianceState = $d.complianceState
                LastSyncDateTime = (Format-IsoDate $d.lastSyncDateTime)
                ManagementAgent = $d.managementAgent; Manufacturer = $d.manufacturer; Model = $d.model
                EnrolledDateTime = (Format-IsoDate $d.enrolledDateTime)
            })
        }
        $out
    }
}

# ---------------------------------------------------------------------------
# EXCHANGE DATASETS
# ---------------------------------------------------------------------------
$script:AcceptedDomainNames = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
$script:ExoMailboxes = @()

if ($script:ExoOk) {

    try {
        $accepted = @(Get-AcceptedDomain)
        foreach ($d in $accepted) { [void]$script:AcceptedDomainNames.Add([string]$d.DomainName) }
    } catch {
        Add-Warning ("AcceptedDomains: " + $_.Exception.Message)
        $accepted = @()
    }

    function Test-ExternalSmtp {
        param([string]$Smtp)
        if ("$Smtp" -eq '') { return $false }
        $addr = $Smtp -replace '^smtp:', ''
        $parts = $addr.Split('@')
        if ($parts.Count -lt 2) { return $false }
        return -not $script:AcceptedDomainNames.Contains($parts[-1])
    }

    Invoke-Dataset 'Mailboxes' {
        $script:ExoMailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName,PrimarySmtpAddress,RecipientTypeDetails,GrantSendOnBehalfTo,ForwardingAddress,ForwardingSmtpAddress,ArchiveStatus,LitigationHoldEnabled)
        $total = $script:ExoMailboxes.Count
        Write-Log "Mailboxes: $total found, collecting statistics ..."
        $out = New-Object System.Collections.Generic.List[object]
        $started = Get-Date
        $i = 0
        foreach ($mbx in $script:ExoMailboxes) {
            $i++
            if (($i % 10) -eq 0 -or $i -eq $total) {
                $elapsed = ((Get-Date) - $started).TotalSeconds
                $remaining = 0
                if ($i -gt 0) { $remaining = ($elapsed / $i) * ($total - $i) }
                Write-Progress -Activity 'Mailbox statistics' -Status "$i of $total - $($mbx.PrimarySmtpAddress)" `
                    -PercentComplete ([int](100 * $i / [math]::Max(1, $total))) -SecondsRemaining ([int]$remaining)
            }
            $sizeGb = ''; $itemCount = ''; $lastAction = ''
            try {
                $stats = Get-EXOMailboxStatistics -Identity $mbx.PrimarySmtpAddress -Properties LastUserActionTime -ErrorAction Stop
                if ($stats) {
                    # robust parse: byte value inside parentheses; separators vary with culture
                    $sizeText = [string]$stats.TotalItemSize
                    if ($sizeText -match '\(([\d,.\s ]+)\s*bytes\)') {
                        $bytes = [double](($Matches[1] -replace '[^\d]', ''))
                        $sizeGb = Format-Num ($bytes / 1GB) 3
                    }
                    $itemCount = Format-Num $stats.ItemCount
                    $lastAction = Format-IsoDate $stats.LastUserActionTime
                }
            } catch {
                Add-Warning ("Mailboxes: statistics failed for {0}: {1}" -f $mbx.PrimarySmtpAddress, $_.Exception.Message)
            }

            $lastReceived = ''; $lastSent = ''
            if ($DeepMailboxScan -and $mbx.RecipientTypeDetails -eq 'UserMailbox') {
                # exact per-folder scan (2 extra calls per mailbox), throttle-aware
                foreach ($try in 1..3) {
                    try {
                        $inbox = @(Get-EXOMailboxFolderStatistics -Identity $mbx.PrimarySmtpAddress -FolderScope Inbox -IncludeOldestAndNewestItems -ErrorAction Stop)
                        if ($inbox.Count -gt 0) { $lastReceived = Format-IsoDate $inbox[0].NewestItemReceivedDate }
                        $sent = @(Get-EXOMailboxFolderStatistics -Identity $mbx.PrimarySmtpAddress -FolderScope SentItems -IncludeOldestAndNewestItems -ErrorAction Stop)
                        if ($sent.Count -gt 0) { $lastSent = Format-IsoDate $sent[0].NewestItemReceivedDate }
                        break
                    } catch {
                        if ($_.Exception.Message -match '(?i)(throttl|429)' -and $try -lt 3) { Start-Sleep -Seconds (10 * $try) }
                        else { break }
                    }
                }
            } elseif (-not $script:ConcealedNames) {
                # fast estimate from the tenant-wide email activity report (D90)
                $key = ([string]$mbx.PrimarySmtpAddress).ToLower()
                if ($script:EmailUserMap.ContainsKey($key)) {
                    $ua = $script:EmailUserMap[$key]
                    if ([string]$ua.SendCount -ne '' -and [double]$ua.SendCount -gt 0) { $lastSent = $ua.LastActivityDate }
                    if ([string]$ua.ReceiveCount -ne '' -and [double]$ua.ReceiveCount -gt 0) { $lastReceived = $ua.LastActivityDate }
                }
            }

            $fwdSmtp = [string]$mbx.ForwardingSmtpAddress
            $fwdAddr = [string]$mbx.ForwardingAddress
            $external = Test-ExternalSmtp $fwdSmtp
            if (-not $external -and $fwdAddr -ne '') {
                try {
                    # ExternalEmailAddress is not in the minimum property set
                    $target = Get-EXORecipient -Identity $fwdAddr -Properties ExternalEmailAddress -ErrorAction Stop
                    if ($target -and $target.PSObject.Properties['ExternalEmailAddress'] -and "$($target.ExternalEmailAddress)" -ne '') {
                        $external = Test-ExternalSmtp ([string]$target.ExternalEmailAddress)
                    } elseif ($target -and $target.RecipientTypeDetails -match 'MailContact|MailUser|GuestMailUser') {
                        $external = $true
                    }
                } catch { }
            }

            $out.Add([ordered]@{
                DisplayName = $mbx.DisplayName; PrimarySmtpAddress = $mbx.PrimarySmtpAddress
                RecipientTypeDetails = [string]$mbx.RecipientTypeDetails
                SizeGB = $sizeGb; ItemCount = $itemCount
                LastUserActionTime = $lastAction
                LastEmailReceived = $lastReceived; LastEmailSent = $lastSent
                ArchiveEnabled = (Format-Bool ($mbx.ArchiveStatus -eq 'Active'))
                LitigationHold = (Format-Bool $mbx.LitigationHoldEnabled)
                ForwardingSmtpAddress = ($fwdSmtp -replace '^smtp:', '')
                ForwardingAddress = $fwdAddr
                ForwardsExternally = (Format-Bool $external)
            })
        }
        Write-Progress -Activity 'Mailbox statistics' -Completed

        # group mailboxes (M365 group mailboxes are not returned by Get-EXOMailbox)
        # -Properties DisplayName: not in Get-EXORecipient's minimum property set
        try {
            $groupMbx = @(Get-EXORecipient -RecipientTypeDetails GroupMailbox -ResultSize Unlimited -Properties DisplayName)
            foreach ($gm in $groupMbx) {
                $out.Add([ordered]@{
                    DisplayName = $gm.DisplayName; PrimarySmtpAddress = $gm.PrimarySmtpAddress
                    RecipientTypeDetails = 'GroupMailbox'
                    SizeGB = ''; ItemCount = ''; LastUserActionTime = ''
                    LastEmailReceived = ''; LastEmailSent = ''
                    ArchiveEnabled = 'FALSE'; LitigationHold = 'FALSE'
                    ForwardingSmtpAddress = ''; ForwardingAddress = ''; ForwardsExternally = 'FALSE'
                })
            }
        } catch {
            Add-Warning ("Mailboxes: group mailboxes not listed: " + $_.Exception.Message)
        }
        $out
    }

    Invoke-Dataset 'MailboxPermissions' {
        # scope: all Shared/Room/Equipment mailboxes + optional extra targets file
        $targets = New-Object System.Collections.Generic.List[object]
        foreach ($mbx in $script:ExoMailboxes) {
            if ('SharedMailbox', 'RoomMailbox', 'EquipmentMailbox' -contains [string]$mbx.RecipientTypeDetails) {
                $targets.Add($mbx)
            }
        }
        $extraFile = Join-Path $script:ScriptDir 'CONFIG\extra_permission_targets.txt'
        if (Test-Path $extraFile) {
            foreach ($line in (Get-Content $extraFile)) {
                $line = $line.Trim()
                if ($line -eq '' -or $line.StartsWith('#')) { continue }
                try {
                    $extra = Get-EXOMailbox -Identity $line -Properties GrantSendOnBehalfTo -ErrorAction Stop
                    $targets.Add($extra)
                } catch { Add-Warning ("MailboxPermissions: extra target '{0}' not found" -f $line) }
            }
        }
        $out = New-Object System.Collections.Generic.List[object]
        $total = $targets.Count; $i = 0
        foreach ($mbx in $targets) {
            $i++
            Write-Progress -Activity 'Mailbox permissions' -Status "$i of $total - $($mbx.PrimarySmtpAddress)" -PercentComplete ([int](100 * $i / [math]::Max(1, $total)))
            # FullAccess (legacy filters: not inherited, no NT AUTHORITY, no SID)
            try {
                $perms = Get-EXOMailboxPermission -Identity $mbx.PrimarySmtpAddress |
                    Where-Object { $_.User -notlike '*NT AUTHO*' -and $_.User -notlike '*S-1-5-21*' -and
                                   -not $_.IsInherited -and $_.AccessRights -contains 'FullAccess' }
                foreach ($p in $perms) {
                    $out.Add([ordered]@{ Mailbox = $mbx.PrimarySmtpAddress; PermissionType = 'FullAccess'; Grantee = [string]$p.User })
                }
            } catch { Add-Warning ("MailboxPermissions FullAccess {0}: {1}" -f $mbx.PrimarySmtpAddress, $_.Exception.Message) }
            # SendAs
            try {
                $rperms = Get-EXORecipientPermission -Identity $mbx.PrimarySmtpAddress |
                    Where-Object { $_.AccessRights -like '*Send*' -and $_.Trustee -notlike '*NT AUTHORITY*' -and
                                   $_.Trustee -notlike '*S-1-5-21*' -and -not $_.IsInherited }
                foreach ($p in $rperms) {
                    $out.Add([ordered]@{ Mailbox = $mbx.PrimarySmtpAddress; PermissionType = 'SendAs'; Grantee = [string]$p.Trustee })
                }
            } catch { Add-Warning ("MailboxPermissions SendAs {0}: {1}" -f $mbx.PrimarySmtpAddress, $_.Exception.Message) }
            # SendOnBehalf
            if ($mbx.GrantSendOnBehalfTo) {
                foreach ($g in @($mbx.GrantSendOnBehalfTo)) {
                    $out.Add([ordered]@{ Mailbox = $mbx.PrimarySmtpAddress; PermissionType = 'SendOnBehalf'; Grantee = [string]$g })
                }
            }
        }
        Write-Progress -Activity 'Mailbox permissions' -Completed
        $out
    }

    $script:DlDirectMembers = @{}
    Invoke-Dataset 'DistributionGroups' {
        $out = New-Object System.Collections.Generic.List[object]
        $dls = @(Get-DistributionGroup -ResultSize Unlimited)
        $ddls = @(Get-DynamicDistributionGroup -ResultSize Unlimited)
        $total = $dls.Count + $ddls.Count; $i = 0
        foreach ($dl in $dls) {
            $i++
            Write-Progress -Activity 'Distribution groups' -Status "$i of $total" -PercentComplete ([int](100 * $i / [math]::Max(1, $total)))
            $members = @()
            try { $members = @(Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -ResultSize Unlimited) }
            catch { Add-Warning ("DistributionGroups: members of {0}: {1}" -f $dl.PrimarySmtpAddress, $_.Exception.Message) }
            $script:DlDirectMembers[[string]$dl.PrimarySmtpAddress] = $members
            $type = 'Distribution'
            if ([string]$dl.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') { $type = 'Mail-enabled security' }
            if ([string]$dl.RecipientTypeDetails -eq 'RoomList') { $type = 'Room list' }
            $out.Add([ordered]@{
                DisplayName = $dl.DisplayName; PrimarySmtpAddress = $dl.PrimarySmtpAddress
                Type = $type; MemberCountDirect = $members.Count
            })
        }
        foreach ($dl in $ddls) {
            $i++
            Write-Progress -Activity 'Distribution groups' -Status "$i of $total" -PercentComplete ([int](100 * $i / [math]::Max(1, $total)))
            $members = @()
            try { $members = @(Get-DynamicDistributionGroupMember -Identity $dl.PrimarySmtpAddress -ResultSize Unlimited) }
            catch { Add-Warning ("DistributionGroups: dynamic members of {0}: {1}" -f $dl.PrimarySmtpAddress, $_.Exception.Message) }
            $script:DlDirectMembers[[string]$dl.PrimarySmtpAddress] = $members
            $out.Add([ordered]@{
                DisplayName = $dl.DisplayName; PrimarySmtpAddress = $dl.PrimarySmtpAddress
                Type = 'Dynamic'; MemberCountDirect = $members.Count
            })
        }
        Write-Progress -Activity 'Distribution groups' -Completed
        $out
    }

    Invoke-Dataset 'DLMembers' {
        # recursive expansion (ported from DLGroupMemberRecursive.ps1) with cycle protection
        function Expand-GroupMember {
            param([string]$Identity, [System.Collections.Generic.HashSet[string]]$Visited)
            if (-not $Visited.Add($Identity.ToLower())) { return @() }   # cycle: already expanded
            $members = $null
            if ($script:DlDirectMembers.ContainsKey($Identity)) {
                $members = $script:DlDirectMembers[$Identity]
            } else {
                try { $members = @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop) }
                catch {
                    try { $members = @(Get-DynamicDistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop) }
                    catch { return @() }
                }
            }
            $flat = New-Object System.Collections.Generic.List[object]
            foreach ($m in $members) {
                if ([string]$m.RecipientType -like '*Group*') {
                    $sub = Expand-GroupMember -Identity ([string]$m.PrimarySmtpAddress) -Visited $Visited
                    foreach ($s in $sub) { $flat.Add($s) }
                } else {
                    $flat.Add($m)
                }
            }
            return $flat
        }
        $out = New-Object System.Collections.Generic.List[object]
        $dlAddresses = @($script:DlDirectMembers.Keys)
        $total = $dlAddresses.Count; $i = 0
        foreach ($addr in $dlAddresses) {
            $i++
            Write-Progress -Activity 'DL recursive members' -Status "$i of $total - $addr" -PercentComplete ([int](100 * $i / [math]::Max(1, $total)))
            $visited = New-Object System.Collections.Generic.HashSet[string]
            $members = Expand-GroupMember -Identity $addr -Visited $visited
            $seen = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($m in $members) {
                $smtp = [string]$m.PrimarySmtpAddress
                if ($smtp -ne '' -and -not $seen.Add($smtp)) { continue }   # dedupe per DL
                $out.Add([ordered]@{
                    DLName = $addr; MemberDisplayName = $m.DisplayName
                    MemberSmtp = $smtp; MemberRecipientType = [string]$m.RecipientType
                })
            }
        }
        Write-Progress -Activity 'DL recursive members' -Completed
        $out
    }

    Invoke-Dataset 'TransportRules' {
        $rules = @(Get-TransportRule -ResultSize Unlimited)
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rules) {
            $out.Add([ordered]@{
                Name = $r.Name; State = [string]$r.State
                Priority = (Format-Num $r.Priority); Comments = [string]$r.Comments
            })
        }
        $out
    }
}

# ---------------------------------------------------------------------------
# Domains (merge EXO accepted domains + Graph domains + DKIM)
# ---------------------------------------------------------------------------
Invoke-Dataset 'Domains' {
    $graphDomains = @{}
    if ($script:GraphOk) {
        try {
            foreach ($d in (Invoke-GraphGet 'v1.0/domains')) { $graphDomains[([string]$d.id).ToLower()] = $d }
        } catch { Add-Warning ("Domains: Graph domains failed: " + $_.Exception.Message) }
    }
    $dkim = @{}
    if ($script:ExoOk) {
        try {
            foreach ($k in @(Get-DkimSigningConfig)) { $dkim[([string]$k.Domain).ToLower()] = [bool]$k.Enabled }
        } catch { Add-Warning ("Domains: DKIM config failed: " + $_.Exception.Message) }
    }
    $out = New-Object System.Collections.Generic.List[object]
    if ($script:ExoOk) {
        foreach ($d in @(Get-AcceptedDomain)) {
            $name = [string]$d.DomainName
            $key = $name.ToLower()
            $pwd = ''; $isDefault = (Format-Bool $d.Default)
            if ($graphDomains.ContainsKey($key)) {
                $gd = $graphDomains[$key]
                if ($null -ne $gd.passwordValidityPeriodInDays) { $pwd = Format-Num $gd.passwordValidityPeriodInDays }
            }
            $dk = 'FALSE'
            if ($dkim.ContainsKey($key)) { $dk = Format-Bool $dkim[$key] }
            $out.Add([ordered]@{
                DomainName = $name; Type = [string]$d.DomainType; IsDefault = $isDefault
                DkimEnabled = $dk; PasswordValidityDays = $pwd
            })
        }
    } else {
        foreach ($key in $graphDomains.Keys) {
            $gd = $graphDomains[$key]
            $pwd = ''
            if ($null -ne $gd.passwordValidityPeriodInDays) { $pwd = Format-Num $gd.passwordValidityPeriodInDays }
            $out.Add([ordered]@{
                DomainName = $gd.id; Type = 'Managed'; IsDefault = (Format-Bool $gd.isDefault)
                DkimEnabled = 'FALSE'; PasswordValidityDays = $pwd
            })
        }
    }
    $out
}

# ---------------------------------------------------------------------------
# Security settings (single row, Graph + EXO)
# ---------------------------------------------------------------------------
Invoke-Dataset 'SecuritySettings' {
    $secDefaults = ''; $caTotal = ''; $caEnabled = ''; $sharing = ''; $smtpDisabled = ''; $auditEnabled = ''
    if ($script:GraphOk) {
        try {
            $sd = Invoke-GraphGet 'v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
            if ($sd.Count -gt 0) { $secDefaults = Format-Bool $sd[0].isEnabled }
        } catch { Add-Warning ("SecuritySettings: security defaults: " + $_.Exception.Message) }
        try {
            $cas = Invoke-GraphGet 'v1.0/identity/conditionalAccess/policies?$select=id,state'
            $caTotal = $cas.Count
            $caEnabled = @($cas | Where-Object { [string]$_.state -eq 'enabled' }).Count
        } catch { Add-Warning ("SecuritySettings: conditional access: " + $_.Exception.Message) }
        try {
            $sp = Invoke-GraphGet 'v1.0/admin/sharepoint/settings'
            if ($sp.Count -gt 0) { $sharing = [string]$sp[0].sharingCapability }
        } catch { Add-Warning ("SecuritySettings: SharePoint tenant settings: " + $_.Exception.Message) }
    }
    if ($script:ExoOk) {
        try {
            $tc = Get-TransportConfig
            $smtpDisabled = Format-Bool $tc.SmtpClientAuthenticationDisabled
        } catch { Add-Warning ("SecuritySettings: transport config: " + $_.Exception.Message) }
        try {
            $oc = Get-OrganizationConfig
            $auditEnabled = Format-Bool (-not $oc.AuditDisabled)
        } catch { Add-Warning ("SecuritySettings: organization config: " + $_.Exception.Message) }
    }
    ,([ordered]@{
        SecurityDefaultsEnabled = $secDefaults; CAPoliciesTotal = $caTotal; CAPoliciesEnabled = $caEnabled
        TenantSharingCapability = $sharing; SmtpAuthDisabled = $smtpDisabled; AuditEnabled = $auditEnabled
    })
}

# ---------------------------------------------------------------------------
# Optional module: external user file access (needs audit role, off by default)
# ---------------------------------------------------------------------------
if ($IncludeExternalFileAccess -and $script:ExoOk) {
    Write-Log "External file access report (optional module) ..."
    $canSearch = $false
    try {
        $null = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -ResultSize 1 -ErrorAction Stop
        $canSearch = $true
    } catch {
        Add-Warning ("ExternalFileAccess: Search-UnifiedAuditLog not available with this account (needs View-Only Audit Logs role). Run it separately with an audit-role account and drop the CSV in MANUAL\ExternalFileAccess.csv. Error: " + $_.Exception.Message)
    }
    if ($canSearch) {
        try {
            $cols = @('Accessed Time','External User','Accessed File','Site URL','File Extension','Workload','More Info')
            $rows = New-Object System.Collections.Generic.List[object]
            $endDate = (Get-Date).Date
            $startDate = $endDate.AddDays(-[math]::Min(89, $ExternalAccessDays))
            $curStart = $startDate
            while ($curStart -lt $endDate) {
                $curEnd = $curStart.AddDays(1)
                if ($curEnd -gt $endDate) { $curEnd = $endDate }
                Write-Progress -Activity 'External file access audit' -Status ("{0:yyyy-MM-dd}" -f $curStart) `
                    -PercentComplete ([int](100 * ($curStart - $startDate).TotalDays / [math]::Max(1, ($endDate - $startDate).TotalDays)))
                # ReturnLargeSet pages by re-invoking the identical call until it
                # drains; the SessionId must be unique per window or the service
                # returns continuation pages of the previous window's query.
                $sessionId = [guid]::NewGuid().ToString('N')
                $pages = 0
                do {
                    $results = @(Search-UnifiedAuditLog -StartDate $curStart -EndDate $curEnd -Operations FileAccessed `
                        -UserIds '*#EXT#*' -SessionId $sessionId -SessionCommand ReturnLargeSet -ResultSize 5000)
                    foreach ($r in $results) {
                        try {
                            $data = $r.AuditData | ConvertFrom-Json
                            $rows.Add([ordered]@{
                                'Accessed Time' = (Format-IsoDate $data.CreationTime)
                                'External User' = $data.UserId
                                'Accessed File' = $data.SourceFileName
                                'Site URL' = $data.SiteUrl
                                'File Extension' = $data.SourceFileExtension
                                'Workload' = $data.Workload
                                'More Info' = ''
                            })
                        } catch { }
                    }
                    $pages++
                    if ($pages -ge 10 -and $results.Count -eq 5000) {
                        Add-Warning ("ExternalFileAccess: more than 50000 events on {0:yyyy-MM-dd}; the rest of that day was skipped." -f $curStart)
                        break
                    }
                } while ($results.Count -eq 5000)
                $curStart = $curEnd
            }
            Write-Progress -Activity 'External file access audit' -Completed
            Write-AuditCsv -FileName 'ExternalFileAccess.csv' -Columns $cols -Rows $rows -Folder $script:ManualPath
            Write-Log ("ExternalFileAccess: {0} events written to MANUAL\ExternalFileAccess.csv" -f $rows.Count)
        } catch {
            Add-Warning ("ExternalFileAccess: " + $_.Exception.Message)
        }
    }
}

# ---------------------------------------------------------------------------
# Wrap up
# ---------------------------------------------------------------------------
$failed = 0
foreach ($k in $script:DatasetStatus.Keys) {
    if ($script:DatasetStatus[$k].status -eq 'FAIL') { $failed++ }
}
$outcome = 'OK'
if ($failed -gt 0) { $outcome = 'PARTIAL' }
Write-RunMeta $outcome

if ($script:ExoOk) {
    try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch { }
}
Write-Host ""
Write-Host ("Done: {0}. {1} datasets, {2} failed, {3} warnings." -f $outcome, $script:DatasetStatus.Count, $failed, $script:Warnings.Count) -ForegroundColor Green
Write-Host ("Data folder: {0}" -f $script:DataPath) -ForegroundColor Green
Write-Host "You can now go back to Excel: the workbook imports the data automatically."
exit 0
