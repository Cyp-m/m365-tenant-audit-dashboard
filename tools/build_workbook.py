"""Build dist/M365_Audit_Dashboard.xlsx from the validated template.

Strategy: byte-level OOXML (zip) surgery on a copy of the template, because the
template's icons are 22 embedded pictures that openpyxl cannot round-trip
(openpyxl parses 0 of them and would silently drop them all on save).
Everything visual (icons, colors, merged cells, both charts, print setup) is
kept as raw parts; only the targeted XML is patched.

What this script does:
 1. Dashboard sheet: replaces every hardcoded KPI with a formula over the D_*
    data sheets, fixes the template's English typos, adds the new lines
    (External forwarding, Conditional Access, Total privileged, MDE discovered,
    findings block), keeping each cell's original style index.
 2. Converts all remaining shared-string label cells to inline strings and
    empties sharedStrings.xml (purges all Goodlife data from the file).
 3. Rebinds chart1 (line, mail traffic) to D_EmailActivity and chart2 (bar,
    Teams activity) to D_TeamsUserActivity by editing the chart XML in place
    (colors/styles untouched).
 4. Adds sheets: Config, RunInfo, 5 M_* manual sheets (visible), 22 hidden D_*
    data sheets with contract headers in row 1 and helper formula columns.
 5. Deletes all 17 legacy data sheets and their parts (tables, query tables,
    pivot cache, connections, customXml, calcChain) and orphaned media.
 6. Sets fullCalcOnLoad so the dashboard recalculates on open.

Run:  python tools/build_workbook.py
Output: dist/M365_Audit_Dashboard.xlsx (next to this repo's dist folder)
"""
import os
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from contract import DATASETS, MANUAL_SHEETS  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.dirname(HERE)
ROOT = os.path.dirname(DIST)
TEMPLATE = os.path.join(ROOT, "reference", "template", "GL_Audit M365 - Dashboard.xlsx")
OUTPUT = os.path.join(DIST, "M365_Audit_Dashboard.xlsx")

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
M = "{%s}" % NS_MAIN

# ---------------------------------------------------------------------------
# Common SKU friendly names, prefilled in the Config pricing table.
# Must stay in sync with $SkuFriendlyNames in Collect-M365Audit.ps1
# (tests/validate.py asserts both lists are identical).
# ---------------------------------------------------------------------------
SKU_MAP = {
    "SPE_E3": "Microsoft 365 E3",
    "SPE_E5": "Microsoft 365 E5",
    "SPE_F1": "Microsoft 365 F3",
    "M365_F1": "Microsoft 365 F1",
    "SPB": "Microsoft 365 Business Premium",
    "O365_BUSINESS_ESSENTIALS": "Microsoft 365 Business Basic",
    "O365_BUSINESS_PREMIUM": "Microsoft 365 Business Standard",
    "SMB_BUSINESS": "Microsoft 365 Apps for business",
    "O365_BUSINESS": "Microsoft 365 Apps for business",
    "OFFICESUBSCRIPTION": "Microsoft 365 Apps for enterprise",
    "STANDARDPACK": "Office 365 E1",
    "ENTERPRISEPACK": "Office 365 E3",
    "ENTERPRISEPREMIUM": "Office 365 E5",
    "ENTERPRISEPREMIUM_NOPSTNCONF": "Office 365 E5 without Audio Conferencing",
    "DESKLESSPACK": "Office 365 F3",
    "EXCHANGESTANDARD": "Exchange Online (Plan 1)",
    "EXCHANGEENTERPRISE": "Exchange Online (Plan 2)",
    "EXCHANGEDESKLESS": "Exchange Online Kiosk",
    "EXCHANGEARCHIVE_ADDON": "Exchange Online Archiving",
    "SHAREPOINTSTANDARD": "SharePoint Online (Plan 1)",
    "SHAREPOINTENTERPRISE": "SharePoint Online (Plan 2)",
    "MCOSTANDARD": "Skype for Business Online (Plan 2)",
    "MCOEV": "Microsoft Teams Phone Standard",
    "MCOMEETADV": "Microsoft 365 Audio Conferencing",
    "MCOPSTN1": "Microsoft Teams Domestic Calling Plan",
    "MCOPSTN2": "Microsoft Teams Domestic and International Calling Plan",
    "PHONESYSTEM_VIRTUALUSER": "Microsoft Teams Phone Resource Account",
    "MEETING_ROOM": "Microsoft Teams Rooms Standard",
    "Microsoft_Teams_Rooms_Pro": "Microsoft Teams Rooms Pro",
    "TEAMS_ESSENTIALS_AAD": "Microsoft Teams Essentials",
    "TEAMS_EXPLORATORY": "Microsoft Teams Exploratory",
    "AAD_PREMIUM": "Microsoft Entra ID P1",
    "AAD_PREMIUM_P2": "Microsoft Entra ID P2",
    "EMS": "Enterprise Mobility + Security E3",
    "EMSPREMIUM": "Enterprise Mobility + Security E5",
    "INTUNE_A": "Microsoft Intune Plan 1",
    "ATP_ENTERPRISE": "Microsoft Defender for Office 365 (Plan 1)",
    "THREAT_INTELLIGENCE": "Microsoft Defender for Office 365 (Plan 2)",
    "DEFENDER_ENDPOINT_P1": "Microsoft Defender for Endpoint P1",
    "WIN_DEF_ATP": "Microsoft Defender for Endpoint P2",
    "ADALLOM_STANDALONE": "Microsoft Defender for Cloud Apps",
    "ATA": "Microsoft Defender for Identity",
    "RIGHTSMANAGEMENT": "Azure Information Protection Premium P1",
    "POWER_BI_STANDARD": "Power BI (free)",
    "POWER_BI_PRO": "Power BI Pro",
    "PBI_PREMIUM_PER_USER": "Power BI Premium Per User",
    "FLOW_FREE": "Power Automate Free",
    "POWERAUTOMATE_ATTENDED_RPA": "Power Automate Premium",
    "POWERAPPS_PER_USER": "Power Apps Premium",
    "PROJECT_P1": "Project Plan 1",
    "PROJECTPROFESSIONAL": "Project Plan 3",
    "PROJECTPREMIUM": "Project Plan 5",
    "VISIO_PLAN1_DEPT": "Visio Plan 1",
    "VISIOCLIENT": "Visio Plan 2",
    "WIN10_PRO_ENT_SUB": "Windows 10/11 Enterprise E3",
    "WIN10_VDA_E5": "Windows 10/11 Enterprise E5",
    "WINDOWS_STORE": "Windows Store for Business",
    "STREAM": "Microsoft Stream",
    "Microsoft_365_Copilot": "Microsoft 365 Copilot",
    "DEVELOPERPACK_E5": "Microsoft 365 E5 Developer",
    "FORMS_PRO": "Dynamics 365 Customer Voice Trial",
    "DYN365_ENTERPRISE_SALES": "Dynamics 365 Sales Enterprise",
    "DYN365_ENTERPRISE_CUSTOMER_SERVICE": "Dynamics 365 Customer Service Enterprise",
    "CCIBOTS_PRIVPREV_VIRAL": "Copilot Studio Viral Trial",
}

# ---------------------------------------------------------------------------
# Dashboard formula map. RefDate = Config!$B$11 (run date, or TODAY() if empty).
# Range sizes come from contract.py capacities.
# ---------------------------------------------------------------------------
REF = "Config!$B$11"
U = "D_Users"; G = "D_Groups"; MB = "D_Mailboxes"; EA = "D_EmailActivity"
SP = "D_SPOSites"; OD = "D_OneDrive"; TM = "D_Teams"; MFA = "D_MFA"
RL = "D_Roles"; LI = "D_Licenses"; DEV_E = "D_DevicesEntra"; DEV_I = "D_DevicesIntune"
APP = "D_EnterpriseApps"; REG = "D_AppRegistrations"; DOM = "D_Domains"
DG = "D_DistributionGroups"; SEC = "D_SecuritySettings"

SHARING_MAP = (
    '=IF({s}!$D$2="","-",IF({s}!$D$2="Disabled","Disabled",'
    'IF({s}!$D$2="ExistingExternalUserSharingOnly","Existing guests",'
    'IF({s}!$D$2="ExternalUserSharingOnly","New and existing guests",'
    'IF({s}!$D$2="ExternalUserAndGuestSharing","Anyone with link",{s}!$D$2)))))'
).format(s=SEC)

MBX_INACTIVE = (
    f'COUNTIF({MB}!$C$2:$C$20001,"UserMailbox")-H15-H16-H17'
)
NEVER_SENT = (
    f'COUNTIFS({MB}!$C$2:$C$20001,"UserMailbox",{MB}!$H$2:$H$20001,"=")'
)
FWD_EXT = f'COUNTIF({MB}!$M$2:$M$20001,TRUE)'
PRIV_NO_MFA = f'COUNTIFS({MFA}!$D$2:$D$20001,TRUE,{MFA}!$B$2:$B$20001,FALSE)'
SECRETS_90 = (
    f'COUNTIFS({REG}!$E$2:$E$3001,">="&{REF},{REG}!$E$2:$E$3001,"<"&({REF}+90))'
)

FORMULAS = {
    # Title
    "C1": '=" DASHBOARD MICROSOFT 365 - "&Config!$B$2',
    # ---- Entra ----
    "A6": f'=COUNTIF({U}!$E$2:$E$20001,"Member")',
    "D6": f'=COUNTIFS({U}!$E$2:$E$20001,"Member",{U}!$F$2:$F$20001,FALSE)',
    "D7": f'=COUNTIFS({U}!$E$2:$E$20001,"Member",{U}!$K$2:$K$20001,TRUE)',
    "D8": f'=COUNTIF({U}!$E$2:$E$20001,"Guest")',
    "C10": (f'=IF(COUNTA({U}!$A$2:$A$20001)=0,"-",'
            f'IF(COUNTIF({U}!$G$2:$G$20001,TRUE)=0,"Cloud-only",'
            f'"Hybrid ("&COUNTIF({U}!$G$2:$G$20001,TRUE)&" synced)"))'),
    # count DisplayName (col B): always present, unlike Id for some group types
    "A14": f'=COUNTA({G}!$B$2:$B$10001)',
    "D14": f'=COUNTIF({G}!$D$2:$D$10001,"Security")',
    "D15": f'=COUNTIF({G}!$D$2:$D$10001,"Microsoft 365")',
    "A19": f'=COUNTA({DEV_E}!$A$2:$A$20001)',
    "A21": f'=COUNTIFS({DEV_E}!$E$2:$E$20001,">="&DATE(YEAR({REF}),1,1))',
    "B21": (f'=COUNTIFS({DEV_E}!$E$2:$E$20001,"<"&({REF}-182))'
            f'+COUNTIFS({DEV_E}!$A$2:$A$20001,"<>",{DEV_E}!$E$2:$E$20001,"=")'),
    "C21": f'=COUNTIF({DEV_I}!$D$2:$D$20001,"noncompliant")',
    "D21": f'=COUNTIFS({DEV_E}!$A$2:$A$20001,"<>",{DEV_E}!$G$2:$G$20001,FALSE)',
    "A27": f'=COUNTIFS({APP}!$A$2:$A$5001,"<>",{APP}!$G$2:$G$5001,FALSE)',
    # ---- Exchange ----
    "F6": f'=COUNTA({MB}!$A$2:$A$20001)',
    "H6": f'=ROUND(SUM({MB}!$D$2:$D$20001),0)&" GB"',
    "G8": (f'=COUNTIF({MB}!$C$2:$C$20001,"RoomMailbox")'
           f'+COUNTIF({MB}!$C$2:$C$20001,"EquipmentMailbox")'),
    "G9": f'=COUNTIF({MB}!$C$2:$C$20001,"SharedMailbox")',
    "G10": f'=COUNTIF({MB}!$C$2:$C$20001,"UserMailbox")',
    "G11": (f'=COUNTIF({MB}!$C$2:$C$20001,"GroupMailbox")'
            f'+COUNTIF({MB}!$C$2:$C$20001,"TeamMailbox")'),
    "G12": f'=COUNTIF({MB}!$C$2:$C$20001,"SchedulingMailbox")',
    "H15": (f'=COUNTIFS({MB}!$C$2:$C$20001,"UserMailbox",'
            f'{MB}!$F$2:$F$20001,">="&({REF}-Config!$B$3))'),
    "H16": (f'=COUNTIFS({MB}!$C$2:$C$20001,"UserMailbox",'
            f'{MB}!$F$2:$F$20001,">="&({REF}-Config!$B$4),'
            f'{MB}!$F$2:$F$20001,"<"&({REF}-Config!$B$3))'),
    "H17": (f'=COUNTIFS({MB}!$C$2:$C$20001,"UserMailbox",'
            f'{MB}!$F$2:$F$20001,">="&({REF}-Config!$B$5),'
            f'{MB}!$F$2:$F$20001,"<"&({REF}-Config!$B$4))'),
    "H18": f'={MBX_INACTIVE}',
    # ROUND(..)&"%" instead of TEXT(..,"0.0%"): TEXT format codes are
    # locale-sensitive and raise #VALUE! on non-English Excel installations.
    "H19": '=IF(G10=0,"-",ROUND((H15+H16)/G10*100,1)&"%")',
    "F30": f'=SUM({EA}!$B$2:$B$401)',
    "G30": f'=SUM({EA}!$C$2:$C$401)',
    "H30": f'=SUM({EA}!$D$2:$D$401)',
    "H32": f'=COUNTA({DG}!$A$2:$A$5001)',
    "H33": (f'=IF(COUNTIF({MB}!$I$2:$I$20001,TRUE)>0,'
            f'"Yes ("&COUNTIF({MB}!$I$2:$I$20001,TRUE)&")","No")'),
    "H34": f'=COUNTA({DOM}!$A$2:$A$301)',
    # ---- SharePoint ----
    "K6": f'=COUNTIFS({SP}!$A$2:$A$10001,"<>",{SP}!$H$2:$H$10001,FALSE)',
    "M8": f'=ROUND(SUMIFS({SP}!$F$2:$F$10001,{SP}!$H$2:$H$10001,FALSE),1)',
    "M9": (f'=COUNTIFS({SP}!$A$2:$A$10001,"<>",{SP}!$H$2:$H$10001,FALSE,'
           f'{SP}!$C$2:$C$10001,"<"&({REF}-Config!$B$5))'
           f'+COUNTIFS({SP}!$A$2:$A$10001,"<>",{SP}!$H$2:$H$10001,FALSE,'
           f'{SP}!$C$2:$C$10001,"=")'),
    "M10": ('=IF(COUNTA(M_SPAdminSites!$A$2:$A$10001)=0,"manual",'
            'COUNTIFS(M_SPAdminSites!$E$2:$E$10001,"<>Team site",'
            'M_SPAdminSites!$I$2:$I$10001,"On"))'),
    "K14": f'=COUNTIF({SP}!$H$2:$H$10001,TRUE)',
    "M16": f'=ROUND(SUMIFS({SP}!$F$2:$F$10001,{SP}!$H$2:$H$10001,TRUE),0)',
    "M17": ('=IF(COUNTA(M_SPAdminSites!$A$2:$A$10001)=0,"manual",'
            'COUNTIFS(M_SPAdminSites!$E$2:$E$10001,"Team site",'
            'M_SPAdminSites!$I$2:$I$10001,"On"))'),
    "M19": (f'=COUNTIFS({SP}!$H$2:$H$10001,TRUE,'
            f'{SP}!$C$2:$C$10001,">="&({REF}-Config!$B$3))'),
    "M20": (f'=COUNTIFS({SP}!$H$2:$H$10001,TRUE,'
            f'{SP}!$C$2:$C$10001,">="&({REF}-Config!$B$4),'
            f'{SP}!$C$2:$C$10001,"<"&({REF}-Config!$B$3))'),
    "M21": (f'=COUNTIFS({SP}!$H$2:$H$10001,TRUE,'
            f'{SP}!$C$2:$C$10001,">="&({REF}-Config!$B$5),'
            f'{SP}!$C$2:$C$10001,"<"&({REF}-Config!$B$4))'),
    "M22": '=K14-M19-M20-M21',
    "M25": f'=SUM({SP}!$D$2:$D$10001)',
    "M27": SHARING_MAP,
    "M29": '=Config!$B$7',
    "M30": '=Config!$B$8',
    "M31": '=Config!$B$9',
    # ---- Teams ----
    "P6": f'=COUNTA({TM}!$A$2:$A$10001)',
    "S6": f'=COUNTIF({TM}!$C$2:$C$10001,"Public")',
    "S7": f'=COUNTIF({TM}!$C$2:$C$10001,"Private")',
    "P10": ('=IF(COUNTA(M_TeamsAdminExport!$A$2:$A$10001)=0,"manual",'
            'SUM(M_TeamsAdminExport!$B$2:$D$10001))'),
    "S10": ('=IF(COUNTA(M_TeamsAdminExport!$A$2:$A$10001)=0,"manual",'
            'SUM(M_TeamsAdminExport!$B$2:$B$10001))'),
    "S11": ('=IF(COUNTA(M_TeamsAdminExport!$A$2:$A$10001)=0,"manual",'
            'SUM(M_TeamsAdminExport!$C$2:$C$10001))'),
    "S12": ('=IF(COUNTA(M_TeamsAdminExport!$A$2:$A$10001)=0,"manual",'
            'SUM(M_TeamsAdminExport!$D$2:$D$10001))'),
    "S13": ('=IF(COUNTA(M_TeamsAdminExport!$A$2:$A$10001)=0,"manual",'
            'SUM(M_TeamsAdminExport!$G$2:$G$10001))'),
    "S15": f'=SUM({TM}!$D$2:$D$10001)',
    "R29": f'=ROUND(SUM({OD}!$F$2:$F$20001),0)',
    "R30": f'=COUNTA({OD}!$A$2:$A$20001)',
    "Q32": (f'=IF(SUM({OD}!$D$2:$D$20001)>=1000000,'
            f'ROUND(SUM({OD}!$D$2:$D$20001)/1000000,1)&"M",'
            f'SUM({OD}!$D$2:$D$20001)&"")'),
    "R34": SHARING_MAP,
    # ---- Security ----
    "U6": f'=COUNTIF({MFA}!$B$2:$B$20001,TRUE)',
    "U9": (f'=COUNTIFS({U}!$E$2:$E$20001,"Member",{U}!$F$2:$F$20001,TRUE,'
           f'{U}!$K$2:$K$20001,TRUE,{U}!$M$2:$M$20001,FALSE)'),
    "W11": (f'=IF({SEC}!$A$2="","-",'
            f'IF({SEC}!$A$2=TRUE,"Enabled","Disabled"))'),
    "W13": (f'=IF(COUNT({DOM}!$E$2:$E$301)=0,"-",'
            f'IF(MIN({DOM}!$E$2:$E$301)>=2147483647,"Never",'
            f'MIN({DOM}!$E$2:$E$301)&" days"))'),
    "X16": f'=COUNTIF({RL}!$A$2:$A$4001,"Global Administrator")',
    "X17": f'=COUNTIF({RL}!$A$2:$A$4001,"SharePoint Administrator")',
    "X18": f'=COUNTIF({RL}!$A$2:$A$4001,"Teams Administrator")',
    "X19": f'=COUNTIF({RL}!$A$2:$A$4001,"User Administrator")',
    "X20": f'=COUNTIF({RL}!$A$2:$A$4001,"Exchange Administrator")',
    "W33": f'=SUMPRODUCT({LI}!$D$2:$D$301,{LI}!$F$2:$F$301)',
    "W34": '=W33*12',
    # ---- Footnotes (dd/MM/yyyy built from DAY/MONTH/YEAR: TEXT date codes
    # are locale-sensitive and raise #VALUE! on non-English Excel) ----
    "A37": (f'="*Very Active: last user action ≤ "&Config!$B$3&" days (since "'
            f'&TEXT(DAY({REF}-Config!$B$3),"00")&"/"'
            f'&TEXT(MONTH({REF}-Config!$B$3),"00")&"/"'
            f'&YEAR({REF}-Config!$B$3)&")"'),
    "A38": '="Active: last user action "&Config!$B$3+1&"-"&Config!$B$4&" days ago"',
    "A39": '="Low Activity: last user action "&Config!$B$4+1&"-"&Config!$B$5&" days ago"',
    "A40": '="Inactive: last user action > "&Config!$B$5&" days ago, or never"',
    "A42": ('="**("&H15&" + "&H16&") / "&G10&" = "'
            '&IF(G10=0,"-",ROUND((H15+H16)/G10*100,1)&"%")&" Healthy Active"'),
}

# Top-12 third-party apps grid (rank k per cell via helper column I on D_EnterpriseApps)
_top_cells = ["A31", "B31", "C31", "D31", "A32", "B32", "C32", "D32",
              "A33", "B33", "C33", "D33"]
for _k, _cell in enumerate(_top_cells, start=1):
    FORMULAS[_cell] = (f'=IFERROR(INDEX({APP}!$A$2:$A$5001,'
                       f'MATCH({_k},{APP}!$I$2:$I$5001,0)),"")')

# License table rows 25..32 (rank helper column G on D_Licenses)
for _r in range(25, 33):
    _k = _r - 24
    FORMULAS[f"U{_r}"] = (f'=IFERROR(INDEX({LI}!$B$2:$B$301,'
                          f'MATCH({_k},{LI}!$G$2:$G$301,0)),"")')
    FORMULAS[f"W{_r}"] = (f'=IF($U{_r}="","",INDEX({LI}!$D$2:$D$301,'
                          f'MATCH({_k},{LI}!$G$2:$G$301,0)))')
    FORMULAS[f"X{_r}"] = (f'=IF($U{_r}="","",'
                          f'IFERROR(VLOOKUP($U{_r},Config!$D$3:$E$103,2,FALSE),0))')

# Label replacements (typo fixes + redefined labels), written as inline strings
LABELS = {
    "B7": "Licensed users",
    "U8": "Not registered",
    "A22": "Seen this\nyear",
    "B22": "Stale > 6\nmonths",
    "D22": "Not Intune\nmanaged",
    "A29": "Applications Top 12",
    "F34": "Number of domains",
}

# Cells to clear (value removed, style kept)
CLEARS = ["A34"]

# New cells: coord -> (kind, content, style_source)
NEW_CELLS = {
    "F35": ("label", "External forwarding", "F34"),
    "H35": ("formula", f'={FWD_EXT}', "H34"),
    "U12": ("label", "Conditional Access", "U11"),
    "W12": ("formula",
            (f'=IF({SEC}!$B$2="","-",'
             f'{SEC}!$B$2&" ("&{SEC}!$C$2&" enabled)")'), "W11"),
    "V21": ("label", "Total privileged", "V20"),
    "X21": ("formula", f'=COUNTA({RL}!$A$2:$A$4001)', "X20"),
    "A25": ("formula",
            ('=IF(COUNTIF(M_MDE_Devices!$X$2:$X$20001,"Can be onboarded")>0,'
             'COUNTIF(M_MDE_Devices!$X$2:$X$20001,"Can be onboarded")'
             '&" devices discovered by MDE (unmanaged)","")'), "A26"),
    # license rows 31-32 are new cells (template stopped at row 30)
    "U31": ("formula", FORMULAS["U31"], "U30"),
    "W31": ("formula", FORMULAS["W31"], "W30"),
    "X31": ("formula", FORMULAS["X31"], "X30"),
    "U32": ("formula", FORMULAS["U32"], "U30"),
    "W32": ("formula", FORMULAS["W32"], "W30"),
    "X32": ("formula", FORMULAS["X32"], "X30"),
}
for _c in ("U31", "W31", "X31", "U32", "W32", "X32"):
    FORMULAS.pop(_c)

# Findings block (red bold, shown only when count > 0)
FINDINGS = {
    "H37": (f'=IF({NEVER_SENT}>0,"Review "&{NEVER_SENT}'
            '&" mailboxes that never sent emails","")'),
    "H38": ('=IF(H18>0,"Disable "&H18&" inactive mailboxes (> "'
            '&Config!$B$5&" days)","")'),
    "H39": '=IF(H35>0,H35&" mailboxes forward email outside the tenant","")',
    "H40": (f'=IF({PRIV_NO_MFA}>0,{PRIV_NO_MFA}'
            '&" privileged accounts without MFA","")'),
    "H41": (f'=IF({SECRETS_90}>0,{SECRETS_90}'
            '&" app secrets expire within 90 days","")'),
}


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------
def col_letter_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def split_coord(coord):
    m = re.match(r"([A-Z]+)(\d+)", coord)
    return m.group(1), int(m.group(2))


def xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def parse_shared_strings(sst_xml):
    """Return the list of strings (concatenating rich-text runs)."""
    root = ET.fromstring(sst_xml)
    out = []
    for si in root.findall(f"{M}si"):
        text = "".join(t.text or "" for t in si.iter(f"{M}t"))
        out.append(text)
    return out


# ---------------------------------------------------------------------------
# styles.xml patching
# ---------------------------------------------------------------------------
def patch_styles(styles_xml, clone_specs):
    """Append fonts/xfs/dxf. clone_specs: list of (key, source_xf_index, font_override).
    Returns (new_xml, {key: new_xf_index}, dxf_index, xf_bold, xf_date, xf_text)."""
    # fonts
    m = re.search(r'<fonts count="(\d+)"', styles_xml)
    font_count = int(m.group(1))
    bold11 = ('<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/>'
              '<family val="2"/><scheme val="minor"/></font>')
    red8 = ('<font><b/><sz val="8"/><color rgb="FFC00000"/><name val="Calibri"/>'
            '<family val="2"/><scheme val="minor"/></font>')
    fid_bold, fid_red = font_count, font_count + 1
    styles_xml = styles_xml.replace(
        f'<fonts count="{font_count}"', f'<fonts count="{font_count + 2}"', 1)
    styles_xml = re.sub(r"</fonts>", bold11 + red8 + "</fonts>", styles_xml, count=1)

    # existing cellXfs (tokenized: xf entries may contain nested <alignment/> etc.)
    m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles_xml, re.S)
    xf_count = int(m.group(1))
    xf_body = m.group(2)
    xf_entries = []
    pos = 0
    while True:
        start = xf_body.find("<xf", pos)
        if start < 0:
            break
        gt = xf_body.index(">", start)
        if xf_body[gt - 1] == "/":
            end = gt + 1
        else:
            end = xf_body.index("</xf>", gt) + len("</xf>")
        xf_entries.append(xf_body[start:end])
        pos = end
    if len(xf_entries) != xf_count:
        raise RuntimeError("cellXfs parse mismatch: %d vs %d" % (len(xf_entries), xf_count))

    new_xfs = []
    idx_map = {}
    next_idx = xf_count

    def add_xf(xml_str, key):
        nonlocal next_idx
        new_xfs.append(xml_str)
        idx_map[key] = next_idx
        next_idx += 1

    # simple utility styles for new sheets
    add_xf('<xf numFmtId="0" fontId="%d" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
           % fid_bold, "__bold")
    add_xf('<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" '
           'applyNumberFormat="1"/>', "__date")
    add_xf('<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" '
           'applyNumberFormat="1"/>', "__text")

    for key, src_idx, font_override in clone_specs:
        src = xf_entries[src_idx]
        if font_override == "red8":
            new = re.sub(r'fontId="\d+"', 'fontId="%d"' % fid_red, src, count=1)
            if 'applyFont=' not in new:
                new = new.replace("<xf ", '<xf applyFont="1" ', 1)
        else:
            new = src
        add_xf(new, key)

    appended = "".join(new_xfs)
    styles_xml = styles_xml.replace(
        f'<cellXfs count="{xf_count}">', f'<cellXfs count="{next_idx}">', 1)
    styles_xml = re.sub(r"</cellXfs>", appended + "</cellXfs>", styles_xml, count=1)

    # dxf for the red conditional format on External forwarding
    m = re.search(r'<dxfs count="(\d+)">', styles_xml)
    if m:
        dxf_count = int(m.group(1))
        dxf = ('<dxf><font><color rgb="FFFFFFFF"/></font>'
               '<fill><patternFill><bgColor rgb="FFC00000"/></patternFill></fill></dxf>')
        styles_xml = styles_xml.replace(
            f'<dxfs count="{dxf_count}">', f'<dxfs count="{dxf_count + 1}">' + dxf, 1)
        dxf_idx = dxf_count
    else:
        dxf_idx = 0
        styles_xml = styles_xml.replace(
            "</cellXfs>",
            "</cellXfs>" , 1)  # dxfs always exists in this template (count=86)
    return styles_xml, idx_map, dxf_idx


# ---------------------------------------------------------------------------
# Dashboard sheet patching (ElementTree)
# ---------------------------------------------------------------------------
def patch_dashboard(sheet_xml, sst, style_idx, dxf_idx):
    ET.register_namespace("", NS_MAIN)
    ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
    ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
    ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
    ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
    ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
    root = ET.fromstring(sheet_xml)
    # mc:Ignorable lists prefixes that MUST be declared on the root; ET drops
    # declarations of unused namespaces, so re-add them as literal attributes.
    root.set("xmlns:xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
    root.set("xmlns:xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
    sheet_data = root.find(f"{M}sheetData")
    rows = {int(r.get("r")): r for r in sheet_data.findall(f"{M}row")}

    def get_cell(coord):
        col, rown = split_coord(coord)
        row = rows.get(rown)
        if row is None:
            return None
        for c in row.findall(f"{M}c"):
            if c.get("r") == coord:
                return c
        return None

    def cell_style(coord):
        c = get_cell(coord)
        return c.get("s", "0") if c is not None else "0"

    def strip_children(c):
        for child in list(c):
            c.remove(child)
        if "t" in c.attrib:
            del c.attrib["t"]

    def set_formula(c, formula):
        strip_children(c)
        f = ET.SubElement(c, f"{M}f")
        f.text = formula[1:] if formula.startswith("=") else formula

    def set_label(c, text):
        strip_children(c)
        c.set("t", "inlineStr")
        is_el = ET.SubElement(c, f"{M}is")
        t = ET.SubElement(is_el, f"{M}t")
        t.text = text
        if text != text.strip() or "\n" in text:
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    def ensure_cell(coord, style):
        col, rown = split_coord(coord)
        row = rows.get(rown)
        if row is None:
            row = ET.Element(f"{M}row", {"r": str(rown), "spans": "1:24"})
            # insert row in order
            all_rows = sheet_data.findall(f"{M}row")
            pos = 0
            for i, rr in enumerate(all_rows):
                if int(rr.get("r")) > rown:
                    break
                pos = i + 1
            sheet_data.insert(pos, row)
            rows[rown] = row
        c = get_cell(coord)
        if c is not None:
            return c
        c = ET.Element(f"{M}c", {"r": coord, "s": style})
        colnum = col_letter_to_num(col)
        pos = 0
        for i, cc in enumerate(row.findall(f"{M}c")):
            ccol, _ = split_coord(cc.get("r"))
            if col_letter_to_num(ccol) > colnum:
                break
            pos = i + 1
        row.insert(pos, c)
        return c

    # 1. formulas over existing cells
    for coord, formula in FORMULAS.items():
        c = get_cell(coord)
        if c is None:
            c = ensure_cell(coord, cell_style_nearby(coord, cell_style))
        set_formula(c, formula)

    # 2. label replacements
    for coord, text in LABELS.items():
        c = get_cell(coord)
        if c is None:
            raise RuntimeError("label target missing: " + coord)
        set_label(c, text)

    # 3. clears
    for coord in CLEARS:
        c = get_cell(coord)
        if c is not None:
            strip_children(c)

    # 4. new cells with cloned styles (force the style even when the template
    #    already had an empty styled cell at that coordinate)
    for coord, (kind, content, key) in NEW_CELLS.items():
        c = ensure_cell(coord, str(style_idx[coord]))
        c.set("s", str(style_idx[coord]))
        if kind == "formula":
            set_formula(c, content)
        else:
            set_label(c, content)

    # 5. findings block (red bold style)
    for coord, formula in FINDINGS.items():
        c = ensure_cell(coord, str(style_idx["__finding"]))
        c.set("s", str(style_idx["__finding"]))
        set_formula(c, formula)

    # 6. convert remaining shared strings to inline strings
    for row in sheet_data.findall(f"{M}row"):
        for c in row.findall(f"{M}c"):
            if c.get("t") == "s":
                v = c.find(f"{M}v")
                text = sst[int(v.text)] if v is not None and v.text else ""
                set_label(c, text)

    # 7. conditional formatting for External forwarding (H35 > 0 -> red)
    cf = ET.Element(f"{M}conditionalFormatting", {"sqref": "H35"})
    rule = ET.SubElement(cf, f"{M}cfRule", {
        "type": "cellIs", "dxfId": str(dxf_idx), "priority": "1", "operator": "greaterThan"})
    f = ET.SubElement(rule, f"{M}formula")
    f.text = "0"
    merge = root.find(f"{M}mergeCells")
    root.insert(list(root).index(merge) + 1, cf)

    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def cell_style_nearby(coord, cell_style_fn):
    """Style for a FORMULAS coordinate that doesn't exist yet: reuse row neighbour."""
    col, rown = split_coord(coord)
    return cell_style_fn(f"{col}{rown - 1}")


# ---------------------------------------------------------------------------
# Chart rebinding (string surgery keeps every visual attribute)
# ---------------------------------------------------------------------------
def patch_chart1(xml_text):
    # series name refs (row 2 header -> row 1 header on D_EmailActivity)
    for col in "BCD":
        xml_text = xml_text.replace(
            f"'Email activity'!${col}$2</c:f>", f"@@EA!${col}$1</c:f>")
        xml_text = xml_text.replace(
            f"'Email activity'!${col}$3:${col}$32</c:f>",
            f"@@EA!${col}$2:${col}$32</c:f>")
    xml_text = xml_text.replace("@@EA", "D_EmailActivity")
    # add categories (ReportDate) to each series, right before <c:val>
    cat = ("<c:cat><c:numRef><c:f>D_EmailActivity!$A$2:$A$32</c:f>"
           "<c:numCache><c:formatCode>dd/mm</c:formatCode><c:ptCount val=\"0\"/>"
           "</c:numCache></c:numRef></c:cat>")
    xml_text = xml_text.replace("<c:val>", cat + "<c:val>")
    return xml_text


def patch_chart2(xml_text):
    src = "TeamsUserActivityCounts10_28_20"
    shift = {"C": "D", "D": "E", "E": "F"}  # PrivateChat, Calls, Meetings
    for old_col, new_col in shift.items():
        xml_text = xml_text.replace(
            f"{src}!${old_col}$1</c:f>", f"@@TU!${new_col}$1</c:f>")
        xml_text = xml_text.replace(
            f"{src}!${old_col}$2:${old_col}$31</c:f>",
            f"@@TU!${new_col}$2:${new_col}$31</c:f>")
    xml_text = xml_text.replace("@@TU", "D_TeamsUserActivity")
    return xml_text


# ---------------------------------------------------------------------------
# New sheet generation
# ---------------------------------------------------------------------------
def _cell_xml(coord, kind, value, style=None):
    s = f' s="{style}"' if style else ""
    if kind == "s":  # inline string
        pres = ' xml:space="preserve"' if (value != value.strip()) else ""
        return (f'<c r="{coord}"{s} t="inlineStr"><is><t{pres}>'
                f"{xml_escape(value)}</t></is></c>")
    if kind == "n":
        return f'<c r="{coord}"{s}><v>{value}</v></c>'
    if kind == "f":
        body = value[1:] if value.startswith("=") else value
        return f'<c r="{coord}"{s}><f>{xml_escape(body)}</f></c>'
    raise ValueError(kind)


def num_to_col(n):
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def sheet_xml(rows_xml, cols_xml="", tab_color=None):
    pr = f'<sheetPr><tabColor rgb="{tab_color}"/></sheetPr>' if tab_color else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS_MAIN}">'
        f"{pr}<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
        '<sheetFormatPr defaultRowHeight="14.5"/>'
        f"{cols_xml}<sheetData>{rows_xml}</sheetData>"
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
        'header="0.3" footer="0.3"/></worksheet>'
    )


def build_data_sheet(ds_key, style_idx):
    """Hidden D_* sheet: headers row 1, helper formula columns where needed."""
    fname, sheet, maxrows, cols = DATASETS[ds_key]
    bold = style_idx["__bold"]
    header_cells = "".join(
        _cell_xml(f"{num_to_col(i + 1)}1", "s", name, bold)
        for i, (name, _t) in enumerate(cols))
    rows = []

    helper_rows = []
    if ds_key == "Users":
        # column M: has MFA registered (join to D_MFA on UPN)
        header_cells += _cell_xml("M1", "s", "HasMfa (helper)", bold)
        for r in range(2, maxrows + 2):
            helper_rows.append((r, _cell_xml(
                f"M{r}", "f",
                f'=IF($A{r}="","",COUNTIFS(D_MFA!$A$2:$A$20001,$C{r},'
                f'D_MFA!$B$2:$B$20001,TRUE)>0)')))
    elif ds_key == "EnterpriseApps":
        # column I: rank of third-party apps by CreatedDateTime (newest = 1)
        header_cells += _cell_xml("I1", "s", "ThirdPartyRank (helper)", bold)
        for r in range(2, maxrows + 2):
            helper_rows.append((r, _cell_xml(
                f"I{r}", "f",
                f'=IF($A{r}="","",IF($G{r}<>FALSE,"",'
                f'COUNTIFS($D$2:$D$5001,">"&$D{r},$G$2:$G$5001,FALSE)'
                f'+COUNTIFS($D$2:$D{r},$D{r},$G$2:$G{r},FALSE)))')))
    elif ds_key == "Licenses":
        # F: unit price from Config pricing, G: rank by Assigned desc (>0 only)
        header_cells += _cell_xml("F1", "s", "UnitPrice (helper)", bold)
        header_cells += _cell_xml("G1", "s", "AssignedRank (helper)", bold)
        for r in range(2, maxrows + 2):
            helper_rows.append((r, _cell_xml(
                f"F{r}", "f",
                f'=IF($B{r}="",0,IFERROR(VLOOKUP($B{r},Config!$D$3:$E$103,2,FALSE),0))')
                + _cell_xml(
                f"G{r}", "f",
                f'=IF($A{r}="","",IF(OR($D{r}="",$D{r}=0),"",'
                f'COUNTIFS($D$2:$D$301,">"&$D{r})+COUNTIF($D$2:$D{r},$D{r})))')))

    rows.append(f'<row r="1">{header_cells}</row>')
    for rown, cells in helper_rows:
        rows.append(f'<row r="{rown}">{cells}</row>')
    return sheet_xml("".join(rows))


def build_manual_sheet(name, style_idx):
    csv_name, headers, instruction = MANUAL_SHEETS[name]
    bold = style_idx["__bold"]
    cells = "".join(
        _cell_xml(f"{num_to_col(i + 1)}1", "s", h, bold) for i, h in enumerate(headers))
    note_col = num_to_col(len(headers) + 2)
    cells += _cell_xml(f"{note_col}1", "s", instruction)
    return sheet_xml(f'<row r="1">{cells}</row>')


def build_config_sheet(style_idx):
    bold, date_s = style_idx["__bold"], style_idx["__date"]
    r = []
    r.append(f'<row r="1">{_cell_xml("A1", "s", "M365 AUDIT - CONFIGURATION", bold)}'
             f'{_cell_xml("D1", "s", "License pricing - type unit price per month", bold)}</row>')
    settings = [
        (2, "Tenant name", "s", "Target Tenant"),
        (3, "Very Active threshold (days)", "n", "30"),
        (4, "Active threshold (days)", "n", "90"),
        (5, "Low Activity threshold (days)", "n", "180"),
        (6, "Currency label", "s", "EUR"),
        (7, "Webpart used (Yes/No) - manual answer", "s", "No"),
        (8, "Workflow used (Yes/No) - manual answer", "s", "No"),
        (9, "Custom dev (Yes/No) - manual answer", "s", "No"),
        (10, "Collect external file access (needs audit role, Yes/No)", "s", "No"),
        (11, "Reference date (run date, automatic)", "f",
         '=IF(RunInfo!$B$4="",TODAY(),RunInfo!$B$4)'),
        (12, "Deep mailbox scan (slow, Yes/No)", "s", "No"),
    ]
    price_rows = {2: _cell_xml("D2", "s", "FriendlyName", bold)
                     + _cell_xml("E2", "s", "UnitPrice/month", bold)}
    for i, friendly in enumerate(sorted(set(SKU_MAP.values()))):
        price_rows[3 + i] = _cell_xml(f"D{3 + i}", "s", friendly)

    for rown, label, kind, val in settings:
        cells = _cell_xml(f"A{rown}", "s", label)
        if kind == "f":
            cells += _cell_xml(f"B{rown}", "f", val, date_s)
        elif kind == "n":
            cells += _cell_xml(f"B{rown}", "n", val)
        else:
            cells += _cell_xml(f"B{rown}", "s", val)
        cells += price_rows.pop(rown, "")
        r.append(f'<row r="{rown}">{cells}</row>')
    for rown in sorted(price_rows):
        r.append(f'<row r="{rown}">{price_rows[rown]}</row>')
    cols = ('<cols><col min="1" max="1" width="45" customWidth="1"/>'
            '<col min="2" max="2" width="16" customWidth="1"/>'
            '<col min="4" max="4" width="45" customWidth="1"/>'
            '<col min="5" max="5" width="16" customWidth="1"/></cols>')
    return sheet_xml("".join(r), cols)


def build_runinfo_sheet(style_idx):
    bold, date_s = style_idx["__bold"], style_idx["__date"]
    r = []
    r.append(f'<row r="1">{_cell_xml("A1", "s", "RUN INFO - filled by the importer", bold)}</row>')
    labels = [(3, "Tenant name"), (4, "Run date"), (5, "Account"), (6, "Tenant id"),
              (7, "Concealed names in usage reports"), (8, "Warnings")]
    for rown, label in labels:
        cells = _cell_xml(f"A{rown}", "s", label)
        if rown == 4:
            cells += f'<c r="B4" s="{date_s}"/>'
        r.append(f'<row r="{rown}">{cells}</row>')
    hdr = (_cell_xml("A10", "s", "Dataset", bold) + _cell_xml("B10", "s", "Rows", bold)
           + _cell_xml("C10", "s", "Status", bold) + _cell_xml("D10", "s", "Message", bold))
    r.append(f'<row r="10">{hdr}</row>')
    cols = ('<cols><col min="1" max="1" width="34" customWidth="1"/>'
            '<col min="2" max="3" width="14" customWidth="1"/>'
            '<col min="4" max="4" width="80" customWidth="1"/></cols>')
    return sheet_xml("".join(r), cols)


# ---------------------------------------------------------------------------
# Main assembly
# ---------------------------------------------------------------------------
def main():
    with zipfile.ZipFile(TEMPLATE) as z:
        parts = {n: z.read(n) for n in z.namelist()}

    sst = parse_shared_strings(parts["xl/sharedStrings.xml"].decode("utf-8"))
    dash_xml = parts["xl/worksheets/sheet1.xml"].decode("utf-8")

    # -- collect source style indices for cloned cells ----------------------
    def style_of(coord):
        m = re.search(r'<c r="%s"(?: s="(\d+)")?' % coord, dash_xml)
        if not m:
            raise RuntimeError("cannot find style source cell " + coord)
        return int(m.group(1) or 0)

    clone_specs = []
    for coord, (_k, _c, src) in NEW_CELLS.items():
        clone_specs.append((coord, style_of(src), None))
    clone_specs.append(("__finding", style_of("H39"), "red8"))

    styles_xml, style_idx, dxf_idx = patch_styles(
        parts["xl/styles.xml"].decode("utf-8"), clone_specs)
    parts["xl/styles.xml"] = styles_xml.encode("utf-8")

    # -- dashboard ----------------------------------------------------------
    parts["xl/worksheets/sheet1.xml"] = patch_dashboard(
        dash_xml, sst, style_idx, dxf_idx)

    # -- purge shared strings (removes all source-tenant data) --------------
    parts["xl/sharedStrings.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="{NS_MAIN}" count="0" uniqueCount="0"/>').encode("utf-8")

    # -- charts --------------------------------------------------------------
    parts["xl/charts/chart1.xml"] = patch_chart1(
        parts["xl/charts/chart1.xml"].decode("utf-8")).encode("utf-8")
    parts["xl/charts/chart2.xml"] = patch_chart2(
        parts["xl/charts/chart2.xml"].decode("utf-8")).encode("utf-8")

    # -- delete legacy parts --------------------------------------------------
    keep_media = set()
    for rel_part in ("xl/drawings/_rels/drawing1.xml.rels",
                     "xl/charts/_rels/chart1.xml.rels",
                     "xl/charts/_rels/chart2.xml.rels"):
        for target in re.findall(r'Target="\.\./media/([^"]+)"', parts[rel_part].decode("utf-8")):
            keep_media.add("xl/media/" + target)

    drop = []
    for name in parts:
        if re.match(r"xl/worksheets/sheet(?!1\.xml)(\d+)\.xml$", name):
            drop.append(name)
        elif re.match(r"xl/worksheets/_rels/sheet(?!1\.xml)(\d+)\.xml\.rels$", name):
            drop.append(name)
        elif name.startswith(("xl/tables/", "xl/queryTables/", "xl/pivotTables/",
                              "xl/pivotCache/", "customXml/")):
            drop.append(name)
        elif name in ("xl/connections.xml", "xl/calcChain.xml",
                      "xl/drawings/drawing2.xml", "xl/drawings/_rels/drawing2.xml.rels"):
            drop.append(name)
        elif name.startswith("xl/media/") and name not in keep_media:
            drop.append(name)
    for name in drop:
        del parts[name]

    # -- new sheets ------------------------------------------------------------
    new_sheets = []  # (sheet name, part name, hidden)
    new_sheets.append(("Config", "xl/worksheets/sheetC1.xml", False))
    parts["xl/worksheets/sheetC1.xml"] = build_config_sheet(style_idx).encode("utf-8")
    new_sheets.append(("RunInfo", "xl/worksheets/sheetC2.xml", False))
    parts["xl/worksheets/sheetC2.xml"] = build_runinfo_sheet(style_idx).encode("utf-8")
    for i, mname in enumerate(MANUAL_SHEETS):
        part = f"xl/worksheets/sheetM{i + 1}.xml"
        parts[part] = build_manual_sheet(mname, style_idx).encode("utf-8")
        new_sheets.append((mname, part, False))
    for i, key in enumerate(DATASETS):
        part = f"xl/worksheets/sheetD{i + 1}.xml"
        parts[part] = build_data_sheet(key, style_idx).encode("utf-8")
        new_sheets.append((DATASETS[key][1], part, True))

    # -- workbook.xml ------------------------------------------------------------
    sheets_xml = ['<sheet name="Dashboard" sheetId="1" r:id="rId1"/>']
    rels_xml = ['<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                'Target="worksheets/sheet1.xml"/>']
    rid = 2
    for name, part, hidden in new_sheets:
        state = ' state="hidden"' if hidden else ""
        sheets_xml.append(
            f'<sheet name="{name}" sheetId="{rid}"{state} r:id="rId{rid}"/>')
        rels_xml.append(
            f'<Relationship Id="rId{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="{part[3:]}"/>')
        rid += 1
    for tgt, typ in (("theme/theme1.xml", "theme"), ("styles.xml", "styles"),
                     ("sharedStrings.xml", "sharedStrings")):
        rels_xml.append(
            f'<Relationship Id="rId{rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{typ}" '
            f'Target="{tgt}"/>')
        rid += 1

    parts["xl/workbook.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{NS_MAIN}" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="30026"/>'
        '<workbookPr codeName="ThisWorkbook" defaultThemeVersion="202300"/>'
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="29040" '
        'windowHeight="15720" tabRatio="812"/></bookViews>'
        f'<sheets>{"".join(sheets_xml)}</sheets>'
        '<calcPr calcId="191029" fullCalcOnLoad="1"/>'
        "</workbook>").encode("utf-8")

    parts["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels_xml) + "</Relationships>").encode("utf-8")

    # -- [Content_Types].xml -------------------------------------------------
    overrides = [
        ("/xl/workbook.xml",
         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ("/xl/worksheets/sheet1.xml",
         "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
    ]
    for _name, part, _h in new_sheets:
        overrides.append(("/" + part,
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))
    overrides += [
        ("/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
        ("/xl/styles.xml",
         "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
        ("/xl/sharedStrings.xml",
         "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"),
        ("/xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml"),
        ("/xl/charts/chart1.xml",
         "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
        ("/xl/charts/style1.xml", "application/vnd.ms-office.chartstyle+xml"),
        ("/xl/charts/colors1.xml", "application/vnd.ms-office.chartcolorstyle+xml"),
        ("/xl/charts/chart2.xml",
         "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
        ("/xl/charts/style2.xml", "application/vnd.ms-office.chartstyle+xml"),
        ("/xl/charts/colors2.xml", "application/vnd.ms-office.chartcolorstyle+xml"),
        ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
        ("/docProps/app.xml",
         "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
    ]
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="bin" '
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>',
          '<Default Extension="png" ContentType="image/png"/>',
          '<Default Extension="rels" '
          'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="svg" ContentType="image/svg+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>']
    for part, ctype in overrides:
        ct.append(f'<Override PartName="{part}" ContentType="{ctype}"/>')
    ct.append("</Types>")
    parts["[Content_Types].xml"] = "".join(ct).encode("utf-8")

    # -- docProps (regenerated, no source-tenant metadata) -----------------------
    parts["docProps/core.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:title>M365 Audit Dashboard</dc:title>"
        "<dc:creator>M365 Audit Toolkit</dc:creator>"
        "</cp:coreProperties>").encode("utf-8")
    parts["docProps/app.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties '
        'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        "<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity>"
        "</Properties>").encode("utf-8")

    # -- write ------------------------------------------------------------------
    os.makedirs(DIST, exist_ok=True)
    if os.path.exists(OUTPUT):
        os.remove(OUTPUT)
    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml must be first for maximum compatibility
        z.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for name, data in parts.items():
            z.writestr(name, data)
    print("written:", OUTPUT)


if __name__ == "__main__":
    main()
