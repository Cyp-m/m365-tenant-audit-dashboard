"""Single source of truth for the DATA\\*.csv contract.

Used by tools/build_workbook.py (to create the D_* sheets) and by
tests/validate.py (to load CSVs exactly like the VBA importer does).

Column types:
  T = text (VBA OpenText FieldInfo code 2)
  D = ISO date yyyy-MM-dd (FieldInfo code 5, xlYMDFormat)
  N = number (FieldInfo code 1, general)
  B = boolean TRUE/FALSE (FieldInfo code 1, general -> Boolean)
"""

# dataset key -> (csv file name, D_ sheet name, max data rows, [(column, type), ...])
DATASETS = {
    "Users": ("Users.csv", "D_Users", 20000, [
        ("Id", "T"), ("DisplayName", "T"), ("UserPrincipalName", "T"), ("Mail", "T"),
        ("UserType", "T"), ("AccountEnabled", "B"), ("OnPremSynced", "B"),
        ("UsageLocation", "T"), ("CreatedDateTime", "D"), ("LicenseCount", "N"), ("IsLicensed", "B"),
    ]),
    "Groups": ("Groups.csv", "D_Groups", 10000, [
        ("Id", "T"), ("DisplayName", "T"), ("Mail", "T"), ("Category", "T"),
        ("Visibility", "T"), ("IsTeam", "B"),
    ]),
    "MFA": ("MFA.csv", "D_MFA", 20000, [
        ("UserPrincipalName", "T"), ("IsMfaRegistered", "B"), ("IsMfaCapable", "B"),
        ("IsAdmin", "B"), ("MethodsRegistered", "T"),
    ]),
    "Roles": ("Roles.csv", "D_Roles", 4000, [
        ("RoleName", "T"), ("MemberDisplayName", "T"), ("MemberUpn", "T"),
        ("MemberType", "T"), ("AssignmentType", "T"),
    ]),
    "EnterpriseApps": ("EnterpriseApps.csv", "D_EnterpriseApps", 5000, [
        ("DisplayName", "T"), ("AppId", "T"), ("PublisherName", "T"), ("CreatedDateTime", "D"),
        ("AccountEnabled", "B"), ("Homepage", "T"), ("IsMicrosoftFirstParty", "B"),
    ]),
    "AppRegistrations": ("AppRegistrations.csv", "D_AppRegistrations", 3000, [
        ("DisplayName", "T"), ("AppId", "T"), ("CreatedDateTime", "D"),
        ("SecretCount", "N"), ("NearestSecretExpiry", "D"), ("CertCount", "N"),
    ]),
    "Licenses": ("Licenses.csv", "D_Licenses", 300, [
        ("SkuPartNumber", "T"), ("FriendlyName", "T"), ("Total", "N"),
        ("Assigned", "N"), ("Available", "N"),
    ]),
    "Mailboxes": ("Mailboxes.csv", "D_Mailboxes", 20000, [
        ("DisplayName", "T"), ("PrimarySmtpAddress", "T"), ("RecipientTypeDetails", "T"),
        ("SizeGB", "N"), ("ItemCount", "N"), ("LastUserActionTime", "D"),
        ("LastEmailReceived", "D"), ("LastEmailSent", "D"), ("ArchiveEnabled", "B"),
        ("LitigationHold", "B"), ("ForwardingSmtpAddress", "T"), ("ForwardingAddress", "T"),
        ("ForwardsExternally", "B"),
    ]),
    "MailboxPermissions": ("MailboxPermissions.csv", "D_MailboxPermissions", 10000, [
        ("Mailbox", "T"), ("PermissionType", "T"), ("Grantee", "T"),
    ]),
    "DistributionGroups": ("DistributionGroups.csv", "D_DistributionGroups", 5000, [
        ("DisplayName", "T"), ("PrimarySmtpAddress", "T"), ("Type", "T"), ("MemberCountDirect", "N"),
    ]),
    "DLMembers": ("DLMembers.csv", "D_DLMembers", 30000, [
        ("DLName", "T"), ("MemberDisplayName", "T"), ("MemberSmtp", "T"), ("MemberRecipientType", "T"),
    ]),
    "Domains": ("Domains.csv", "D_Domains", 300, [
        ("DomainName", "T"), ("Type", "T"), ("IsDefault", "B"),
        ("DkimEnabled", "B"), ("PasswordValidityDays", "N"),
    ]),
    "TransportRules": ("TransportRules.csv", "D_TransportRules", 1000, [
        ("Name", "T"), ("State", "T"), ("Priority", "N"), ("Comments", "T"),
    ]),
    "EmailActivity": ("EmailActivity.csv", "D_EmailActivity", 400, [
        ("ReportDate", "D"), ("Send", "N"), ("Receive", "N"), ("Read", "N"),
    ]),
    "EmailUserActivity": ("EmailUserActivity.csv", "D_EmailUserActivity", 20000, [
        ("UserPrincipalName", "T"), ("LastActivityDate", "D"), ("SendCount", "N"),
        ("ReceiveCount", "N"), ("ReadCount", "N"),
    ]),
    "SPOSites": ("SPOSites.csv", "D_SPOSites", 10000, [
        ("SiteUrl", "T"), ("OwnerDisplayName", "T"), ("LastActivityDate", "D"),
        ("FileCount", "N"), ("ActiveFileCount", "N"), ("StorageUsedGB", "N"),
        ("RootWebTemplate", "T"), ("IsTeamsConnected", "B"),
    ]),
    "OneDrive": ("OneDrive.csv", "D_OneDrive", 20000, [
        ("OwnerUpn", "T"), ("OwnerDisplayName", "T"), ("LastActivityDate", "D"),
        ("FileCount", "N"), ("ActiveFileCount", "N"), ("StorageUsedGB", "N"),
    ]),
    "Teams": ("Teams.csv", "D_Teams", 10000, [
        ("TeamName", "T"), ("GroupId", "T"), ("Visibility", "T"), ("MemberCount", "N"),
        ("ActiveUsers90d", "N"), ("ChannelMessages90d", "N"), ("LastActivityDate", "D"),
    ]),
    "TeamsUserActivity": ("TeamsUserActivity.csv", "D_TeamsUserActivity", 20000, [
        ("UserPrincipalName", "T"), ("LastActivityDate", "D"), ("TeamChatMessages", "N"),
        ("PrivateChatMessages", "N"), ("Calls", "N"), ("Meetings", "N"),
    ]),
    "DevicesEntra": ("DevicesEntra.csv", "D_DevicesEntra", 20000, [
        ("DisplayName", "T"), ("OS", "T"), ("OSVersion", "T"), ("TrustType", "T"),
        ("LastSignIn", "D"), ("IsCompliant", "B"), ("IsManaged", "B"), ("RegisteredDateTime", "D"),
    ]),
    "DevicesIntune": ("DevicesIntune.csv", "D_DevicesIntune", 20000, [
        ("DeviceName", "T"), ("OS", "T"), ("OSVersion", "T"), ("ComplianceState", "T"),
        ("LastSyncDateTime", "D"), ("ManagementAgent", "T"), ("Manufacturer", "T"),
        ("Model", "T"), ("EnrolledDateTime", "D"),
    ]),
    "SecuritySettings": ("SecuritySettings.csv", "D_SecuritySettings", 2, [
        ("SecurityDefaultsEnabled", "B"), ("CAPoliciesTotal", "N"), ("CAPoliciesEnabled", "N"),
        ("TenantSharingCapability", "T"), ("SmtpAuthDisabled", "B"), ("AuditEnabled", "B"),
    ]),
}

# Manual landing sheets: sheet name -> (MANUAL csv file, headers, instruction)
MANUAL_SHEETS = {
    "M_TeamsAdminExport": ("TeamsAdminExport.csv", [
        "Name", "Standard Channels", "Private Channels", "Shared Channels", "Team Members",
        "Owners", "Guests", "Privacy", "Status", "Classification", "Groups Id",
        "Expiration Date", "Description", "Sensitivity Label"],
        "MANUAL INPUT - optional. Export from Teams admin center > Teams > Manage teams > Export. "
        "Save as TeamsAdminExport.csv in the MANUAL folder, or paste rows below."),
    "M_SPAdminSites": ("SPAdminSites.csv", [
        "Site name", "URL", "Storage used (GB)", "Hub", "Template", "Last activity (UTC)",
        "Created by", "Files", "External sharing"],
        "MANUAL INPUT - optional. Export from SharePoint admin center > Active sites > Export. "
        "Save as SPAdminSites.csv in the MANUAL folder, or paste rows below."),
    "M_MDE_Devices": ("MDE_Devices.csv", [
        "Device ID", "Device Name", "Device Category", "Device Type", "Device Subtype",
        "Discovery sources", "Domain", "AAD Device Id", "First Seen", "Last device update",
        "OS Platform", "OS Distribution", "OS Version", "OS Build", "Windows 10 Version",
        "Tags", "Group", "Is AAD Joined", "Device IPs", "Device MACs", "Risk Level",
        "Exposure Level", "Health Status", "Onboarding Status", "Device Role", "Cloud Platforms",
        "Is Internet Facing", "Enrollment Status Code", "Managed By", "Enrollment Status",
        "Vendor", "Model"],
        "MANUAL INPUT - optional. Export from security.microsoft.com > Assets > Devices > Export. "
        "Save as MDE_Devices.csv in the MANUAL folder, or paste rows below."),
    "M_ExternalFileAccess": ("ExternalFileAccess.csv", [
        "Accessed Time", "External User", "Accessed File", "Site URL", "File Extension",
        "Workload", "More Info"],
        "MANUAL INPUT - optional. Needs an account with an audit role (View-Only Audit Logs). "
        "Run Collect-M365Audit.ps1 with -IncludeExternalFileAccess, or drop ExternalFileAccess.csv "
        "in the MANUAL folder."),
    "M_ITCosts": ("ITCosts.csv", [
        "Cost item", "Category", "Monthly cost", "Yearly cost", "Currency", "Comment"],
        "MANUAL INPUT - optional. Free table for local IT costs (POS, SAP, firewall, internet...). "
        "Type rows directly or drop ITCosts.csv in the MANUAL folder."),
}

# FieldInfo code used by the VBA importer / Refresh-Workbook.ps1 for each type
FIELDINFO = {"T": 2, "D": 5, "N": 1, "B": 1}
