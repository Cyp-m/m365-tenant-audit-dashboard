# Design decisions

Every judgment call made while building the tool, with a one-line rationale.

## Workbook construction

1. **Byte-level OOXML surgery instead of an openpyxl round-trip.** The
   template's icons are 22 embedded pictures (PNG+SVG pairs) that openpyxl
   parses as zero images and would silently drop on save. The builder patches
   the raw XML parts of a byte-copy, so every visual (icons, colors, merges,
   charts, print setup) is preserved exactly.
2. **All Goodlife data purged from the shipped workbook.** The 17 legacy data
   sheets are deleted, `sharedStrings.xml` is emptied (labels become inline
   strings), document properties are regenerated, and the OneDrive absolute
   path is removed. A carve-in tool must not leak the previous target's data.
3. **Legacy query artifacts removed.** Tables, query tables, the pivot cache,
   `connections.xml` and `customXml` belonged to the deleted sheets; keeping
   them would leave phantom data connections.
4. **Charts rebound by editing the original chart XML strings**, not by
   rebuilding chart objects - this keeps the exact line/bar styling. Cached
   values inside the chart XML stay stale; Excel refreshes them on open
   (`fullCalcOnLoad` is set).
5. **Line chart gets ReportDate as category axis** (spec). The bar chart keeps
   no category axis, like the template.
6. **The Teams bar chart now shows the 30 most active users** (rows 2-31 of
   D_TeamsUserActivity, sorted descending by the collector). The template
   charted daily counts, but the CSV contract carries the per-user report;
   sorting makes the fixed 30-row window meaningful.
7. **Bar chart legend shows the contract column names** (PrivateChatMessages,
   Calls, Meetings) because series names point at the D_ sheet headers.
   Accepted cosmetic trade-off for header-assertable imports.
8. **Locale-proof formulas.** `TEXT()` date/percent format codes raise
   `#VALUE!` on non-English Excel (verified on this French installation), so
   dates are built from DAY/MONTH/YEAR + TEXT(n,"00") and percentages from
   `ROUND(x*100,1)&"%"`.
9. **Fixed generous ranges per dataset** (e.g. D_Users rows 2..20001), sized
   for a mid-size tenant: Users/MFA/Mailboxes/EmailUserActivity/OneDrive/
   TeamsUserActivity/devices 20 000; DLMembers 30 000; Groups/SPOSites/Teams
   10 000; MailboxPermissions 10 000; EnterpriseApps 5 000;
   DistributionGroups 5 000; Roles 4 000; AppRegistrations 3 000;
   TransportRules 1 000; EmailActivity 400; Domains/Licenses 300. The importer
   caps at these sizes and reports a warning if a file is bigger.
10. **Helper formula columns live on the D_ sheets** (D_Users!M HasMfa,
    D_EnterpriseApps!I third-party rank, D_Licenses!F price + G rank) and the
    importers clear only the contract columns, so helpers survive every
    import. Headers of helper columns carry the "(helper)" suffix.
11. **Reference date** = run date stamped on RunInfo!B4 by the importer;
    Config!B11 falls back to `TODAY()` when the workbook has never been
    filled (single volatile cell, accepted).
12. **New red-bold style for findings** (FFC00000 on the footnote band fill)
    cloned from the template's footnote style; a `dxf` + conditional format
    turns the External forwarding count red when > 0.

## Dashboard semantics

13. **Groups total counts DisplayName** (column B), not Id - Id can be empty
    for synthesized rows and DisplayName is always present.
14. **Groups "Security" counts pure security groups only**; mail-enabled
    security groups exist in the data (Category "Mail-enabled security") but
    are not a dashboard line, same as the template.
15. **Mailboxes.csv contains only real mailboxes**: UserMailbox,
    SharedMailbox, RoomMailbox, EquipmentMailbox, SchedulingMailbox,
    GroupMailbox. Room line = Room + Equipment; Groups line = GroupMailbox +
    TeamMailbox; Booking = SchedulingMailbox. The five lines always sum to
    the total (asserted by tests).
16. **Group mailboxes carry no size/statistics** (Get-EXOMailbox does not
    return them; the legacy dashboard did not size them either).
17. **Activity bands apply to UserMailbox only** (like the template) and
    Inactive is computed as the remainder, so the four bands always sum to
    the UserMailbox count; mailboxes with no LastUserActionTime land in
    Inactive.
18. **Band thresholds** are Config values: Very Active <= 30 days, Active <=
    90, Low Activity <= 180, Inactive beyond (or never). The footnotes are
    generated from these values and the run date.
19. **Health Score** = (VeryActive + Active) / UserMailbox count, shown with
    one decimal, "-" when there are no user mailboxes.
20. **LastEmailSent/Received default to an estimate**: the tenant-wide
    getEmailActivityUserDetail(D90) report (one Graph call). A user with
    SendCount > 0 gets LastEmailSent = report LastActivityDate. Exact
    per-folder dates only with `-DeepMailboxScan` (the slow legacy behavior,
    kept throttle-aware). The "never sent emails" finding counts UserMailbox
    rows with an empty LastEmailSent - in fast mode this means "no send
    activity in the last 90 days and none known".
21. **ForwardsExternally** = ForwardingSmtpAddress domain not in the accepted
    domains, or ForwardingAddress resolving to a MailContact/MailUser/guest
    with an external address. Displayed as a new Exchange line with a red
    highlight when > 0.
22. **Mailbox permissions scope** = all Shared/Room/Equipment mailboxes
    automatically, plus any addresses in `CONFIG\extra_permission_targets.txt`
    next to the collector. The legacy filters are kept: not inherited, no
    `NT AUTHORITY*`, no `S-1-5-21*` SIDs.
23. **DL members are expanded recursively with cycle protection** (visited
    set per root DL) and per-DL deduplication; dynamic distribution groups
    use the Get-DynamicDistributionGroupMember fallback, as in the legacy
    script.
24. **SharePoint "standalone" = usage-report Root Web Template <> "Group"**;
    the SP pillar's first block is standalone sites, the second is
    teams-connected sites. Inactive sites = standalone with no activity date
    or older than the Low Activity threshold.
25. **Per-site external sharing** comes from the manual SP admin export:
    standalone line counts rows with Template <> "Team site" and sharing
    "On"; the Teams-site line counts Template = "Team site". "manual" is
    displayed until the export is provided.
26. **Tenant-level sharing labels**: Disabled / "Existing guests" /
    "New and existing guests" / "Anyone with link" mapped from
    sharingCapability; unknown values shown raw.
27. **SPO usage report may hide site URLs** (newer tenants): the site id is
    used as SiteUrl and a warning is logged.
28. **OneDrive rows flagged "Is Deleted" are skipped.**
29. **Devices strip redefined with formula-friendly meanings**: Seen this
    year (Entra lastSignIn >= Jan 1 of the run year), Stale > 6 months
    (lastSignIn older than 182 days **or never seen**), Non-compliant
    (Intune complianceState = "noncompliant"), Not Intune-managed (Entra
    isManaged = FALSE; the collector writes FALSE when Graph returns null).
30. **MDE line** (new, below the strip) counts MDE export rows with
    Onboarding Status = "Can be onboarded" - the discovered-but-unmanaged
    population.
31. **Enterprise applications** = service principals of type Application
    whose appOwnerOrganizationId is not one of the two well-known Microsoft
    tenant ids (f8cdef31-..., 72f988bf-...). "Applications Top 10" was
    relabeled **Top 12** and fills a 4x3 grid with the 12 most recently
    created third-party apps (the template's 13th slot is cleared).
32. **"Not registered" (MFA)** = enabled, licensed members without MFA
    registration - replaces the template's "Pending activation", which had no
    API-backed meaning.
33. **Conditional Access line added** (U12): "n (m enabled)". Empty tenant
    shows "-".
34. **Role lines count assignments** (rows in D_Roles, active + PIM-eligible),
    not distinct people - same reading as the template's numbers. "Total
    privileged" = all directory role assignments collected.
35. **PIM-eligible assignments are best-effort** (needs Entra P2); failure
    logs a warning and continues.
36. **Password expire** = MIN(passwordValidityPeriodInDays) over domains that
    expose it; >= 2147483647 means "Never"; "-" when Graph domains were not
    readable.
37. **App secret finding** counts secrets expiring in [run date, run date +
    90 days); already-expired secrets are not counted ("expire within 90
    days" means the future).
38. **Licenses block**: 8 visible rows (template had 6; rows 31-32 reuse the
    same styles) showing SKUs with Assigned > 0 ranked by Assigned. The
    monthly/yearly totals are SUMPRODUCT over **all** SKUs in D_Licenses (not
    just the 8 visible) so costs are never undercounted. Unit prices come
    from the Config pricing table by FriendlyName (VLOOKUP, 0 when blank or
    missing). Total = prepaidUnits.enabled, Available = Total - consumedUnits.
39. **SKU friendly names**: 66 common SKUs embedded (subset of Microsoft's
    licensing reference), falling back to the raw skuPartNumber. The same
    table exists in the collector and the builder; the test suite asserts
    both copies are identical. The Config pricing table is prefilled with the
    66 friendly names.
40. **Currency**: the € number formats of the template are kept as-is; the
    Config "Currency label" documents what the typed prices mean. Excel
    number formats cannot reference cells, and the template's presentation is
    validated - so the label is informational.
41. **Webpart / Workflow / Custom dev stay manual Yes/No answers** on Config
    (not detectable read-only), referenced by the dashboard.

## Collector

42. **Custom CSV writer** (UTF-8 BOM, comma, every field quoted, embedded
    newlines replaced by spaces, invariant-culture numbers and ISO dates).
    `Export-Csv` formats decimals with the machine culture ("6,76" on a
    French PC), which would corrupt the contract.
43. **Graph access via Invoke-MgGraphRequest + Microsoft.Graph.Authentication
    only.** No other Graph sub-modules are loaded - the spec asks for fast
    module load, and raw REST paths make the collected fields explicit.
44. **429 throttling**: exponential backoff (5s..60s, 5 retries) on top of
    the SDK's built-in retry; mailbox folder statistics (deep scan) retry
    with 10s/20s waits.
45. **403 responses are mapped to the missing delegated scope** (per-endpoint
    hint table) so the target admin knows exactly what to consent to.
46. **Every dataset runs in its own try/catch**; on failure a header-only CSV
    is written so the importer clears stale data instead of mixing runs, the
    error is logged, recorded in run.json and in datasets_status.csv.
47. **datasets_status.csv** (Dataset, Rows, Status, Message, Seconds) is
    written next to run.json so VBA can fill RunInfo without a JSON parser;
    run.json also carries `warningsCount` for the same reason.
48. **_DONE.flag content is `OK` / `PARTIAL` / `FAILED`** + timestamp; the
    flag is written even on failure so the Excel wait loop always terminates.
49. **TotalItemSize parsed by regex** on the byte value inside parentheses,
    stripping every non-digit (group separators differ across cultures); the
    legacy split-on-parenthesis was fragile.
50. **Teams member counts** via `GET /groups/{id}/members/$count`
    (ConsistencyLevel: eventual), one call per team, tolerated to fail per
    team.
51. **EmailUserActivity is collected before mailboxes** so the mailbox loop
    can consume the per-user activity map; when the tenant conceals report
    names the estimate is skipped (warning explains why).
52. **Roles**: activated directory roles + members, then PIM eligibility
    schedules with a roleDefinition id->name map.
53. **Domains** merge EXO accepted domains (authoritative list, Type,
    Default) with Graph domains (passwordValidityPeriodInDays) and DKIM
    state; Graph-only fallback when -SkipExchange.
54. **External file access module is off by default**, probes
    Search-UnifiedAuditLog first, then ports the o365reports logic (daily
    windows, FileAccessed, UserIds *#EXT#*) and writes
    `MANUAL\ExternalFileAccess.csv` so it lands on the M_ sheet like a manual
    drop.
55. **Sample mode (-SampleData)** requires the ImportExcel module (auto
    installed) to read the xlsx exports without Excel. Everything synthesized
    is deterministic (MFA: every 3rd user unregistered; Intune compliance:
    every 10th device noncompliant; MDE: every 8th device duplicated as
    "Can be onboarded"; security settings fixed row) - no randomness, so test
    assertions are stable. Missing exports (Groups, MFA, EmailUserActivity,
    security settings, Entra devices) are synthesized from the closest
    available sample and flagged in run.json warnings.
56. **Sample mode anchors the run date to the data** (max LastUserActionTime
    + 5 days) so the activity bands, footnotes and findings are exercised
    with non-zero values.

## Importers (VBA + macro-free)

57. **Workbooks.OpenText with an explicit FieldInfo per column**
    (text/YMD-date/general), Origin 65001 (UTF-8), comma-only, Local:=False -
    a French Excel cannot mis-parse dates, decimals or booleans.
58. **Header assertion before every import**: any mismatch aborts that file
    with a message naming the column, the got/expected values and the likely
    cause (collector/workbook version drift). Manual sheets are asserted
    too, except the free-format M_ITCosts.
59. **Import clears only the contract columns** from row 2 down to the
    capacity row, then pastes values - helper columns and everything else
    survive; sheets are never deleted or recreated (the charts point at fixed
    ranges).
60. **run.json is read via ADODB.Stream (UTF-8)** and parsed with a minimal
    string scanner - no external VBA references, per spec.
61. **modImport.ImportAllSilent** entry point suppresses all message boxes
    for automation and tests.
62. **RunAudit launches a visible PowerShell window** on purpose (browser
    sign-in + progress), prefers pwsh.exe, falls back to powershell.exe,
    waits on `DATA\_DONE.flag` with a 120-minute timeout, ESC cancels the
    wait without killing the collector.
63. **Refresh-Workbook.ps1 opens Excel with AutomationSecurity =
    ForceDisable** so macros never run in the macro-free path, and routes
    every cell write through a Set-CellValue helper: PowerShell 5.1's COM
    binder intermittently misbinds chained parameterized-property setters
    (`.Cells.Item(r,c).Value2 = v` failed with "cannot cast Int32/Double to
    String" - reproduced during validation).
64. **The dataset contract exists in three places** (contract.py, modImport.bas,
    Refresh-Workbook.ps1) because each runtime is self-contained;
    tests/validate.py parses all three and fails if they ever diverge.

## Validation

65. **The test suite drives the real pipeline**: real collector (sample
    mode), the real macro-free importer, real Excel recalculation, then
    openpyxl + pandas assertions - 70+ KPI checks, zero-formula-error scans
    on sample AND empty data, a one-page PDF export check, chart binding
    checks and the cross-implementation consistency checks above.
66. **Excel-vs-Python rounding**: Excel ROUND is half-away-from-zero, Python
    round() is banker's - the tests use a half-up helper to compare.
67. **The end-to-end VBA path was also executed** during development (real
    xlsm, ImportAllSilent via COM) and produced the same dashboard as the
    macro-free path.

## Post-review fixes (adversarial review, 2026-07-08)

68. **Search-UnifiedAuditLog paging rewritten**: the first port made one call
    per daily window with a reused fixed SessionId - ReturnLargeSet pages by
    re-invoking the identical call until it drains, and a reused SessionId
    across different windows returns continuation pages of the previous
    window's query (silent loss/duplication). Now: fresh GUID SessionId per
    window + do/while loop until a partial page, capped at 10 pages (50 000
    events) per day with a warning, matching the legacy script's intent.
69. **Get-EXORecipient property sets**: DisplayName (group mailbox listing)
    and ExternalEmailAddress (forwarding-target check) are NOT in the EXO
    REST minimum property set and were silently null - both calls now request
    them explicitly with -Properties.
70. **$top removed from userRegistrationDetails and managedDevices** - these
    endpoints do not reliably accept large page sizes; @odata.nextLink paging
    covers them anyway.
71. **Sample mode degrades gracefully too**: top-level sample file loads
    (DLGroupMember, Intune inventory, MANUAL exports) are individually
    guarded, so a missing sample file fails only its datasets (verified with
    a partial sample folder: 554 mailboxes collected while 9 datasets FAIL
    with header-only CSVs and the run still ends OK).
72. **OneDrive/SharePoint-opened workbooks are supported** (found by the user
    on first click: run-time error 52). With AutoSave/OneDrive,
    `ThisWorkbook.Path` returns an https:// URL and every VBA file function
    fails. `modImport.MapUrlToLocal` maps the URL back to the local synced
    folder (OneDrive environment variables + %USERPROFILE%\OneDrive* roots,
    longest-matching URL tail that exists on disk, %20 decoding); both
    RunAudit and ImportAll use it, and show a plain-English message when no
    local folder exists instead of error 52. Unit-tested via COM with the
    real tenant URL, an encoded URL, a local path and an unknown URL.
