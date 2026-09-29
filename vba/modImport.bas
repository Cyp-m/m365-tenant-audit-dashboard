Attribute VB_Name = "modImport"
Option Explicit

' ===========================================================================
' modImport - loads DATA\*.csv into the hidden D_* sheets, MANUAL\*.csv into
' the M_* sheets, and stamps the RunInfo sheet from run.json.
'
' The CSV contract (file names, sheet names, column count, column types) must
' match Collect-M365Audit.ps1 and the workbook built by tools/build_workbook.py.
' Types: T = text, D = ISO date (yyyy-MM-dd), N = number, B = boolean.
' ===========================================================================

Private Const MAX_LOG_ROWS As Long = 40

' When True, no message boxes are shown (used by automation and tests).
Public SilentMode As Boolean

' "csv file|target sheet|max data rows|type string (one letter per column)"
Private Function DatasetTable() As Variant
    DatasetTable = Array( _
        "Users.csv|D_Users|20000|TTTTTBBTDNB", _
        "Groups.csv|D_Groups|10000|TTTTTB", _
        "MFA.csv|D_MFA|20000|TBBBT", _
        "Roles.csv|D_Roles|4000|TTTTT", _
        "EnterpriseApps.csv|D_EnterpriseApps|5000|TTTDBTB", _
        "AppRegistrations.csv|D_AppRegistrations|3000|TTDNDN", _
        "Licenses.csv|D_Licenses|300|TTNNN", _
        "Mailboxes.csv|D_Mailboxes|20000|TTTNNDDDBBTTB", _
        "MailboxPermissions.csv|D_MailboxPermissions|10000|TTT", _
        "DistributionGroups.csv|D_DistributionGroups|5000|TTTN", _
        "DLMembers.csv|D_DLMembers|30000|TTTT", _
        "Domains.csv|D_Domains|300|TTBBN", _
        "TransportRules.csv|D_TransportRules|1000|TTNT", _
        "EmailActivity.csv|D_EmailActivity|400|DNNN", _
        "EmailUserActivity.csv|D_EmailUserActivity|20000|TDNNN", _
        "SPOSites.csv|D_SPOSites|10000|TTDNNNTB", _
        "OneDrive.csv|D_OneDrive|20000|TTDNNN", _
        "Teams.csv|D_Teams|10000|TTTNNND", _
        "TeamsUserActivity.csv|D_TeamsUserActivity|20000|TDNNNN", _
        "DevicesEntra.csv|D_DevicesEntra|20000|TTTTDBBD", _
        "DevicesIntune.csv|D_DevicesIntune|20000|TTTTDTTTD", _
        "SecuritySettings.csv|D_SecuritySettings|2|BNNTBB")
End Function

' "csv file|target sheet|max data rows|assert headers (1/0)"
Private Function ManualTable() As Variant
    ManualTable = Array( _
        "TeamsAdminExport.csv|M_TeamsAdminExport|10000|1", _
        "SPAdminSites.csv|M_SPAdminSites|10000|1", _
        "MDE_Devices.csv|M_MDE_Devices|20000|1", _
        "ExternalFileAccess.csv|M_ExternalFileAccess|50000|1", _
        "ITCosts.csv|M_ITCosts|1000|0")
End Function

' ---------------------------------------------------------------------------
' Workbook folder as a LOCAL path.
' When the workbook is opened from a OneDrive/SharePoint synced folder,
' ThisWorkbook.Path returns an https:// URL; file functions (Dir, Kill, Open)
' then fail with run-time error 52. This maps the URL back to the local
' synced folder.
' ---------------------------------------------------------------------------
Public Function WorkbookLocalPath() As String
    WorkbookLocalPath = MapUrlToLocal(ThisWorkbook.Path)
End Function

Public Function MapUrlToLocal(ByVal p As String) As String
    If InStr(1, p, "http", vbTextCompare) <> 1 Then
        MapUrlToLocal = p            ' already a local path (or empty)
        Exit Function
    End If

    ' candidate local OneDrive roots
    Dim roots(9) As String, nRoots As Long
    Dim v As Variant
    For Each v In Array(Environ$("OneDriveCommercial"), Environ$("OneDrive"), Environ$("OneDriveConsumer"))
        If CStr(v) <> "" And nRoots <= 6 Then
            roots(nRoots) = CStr(v)
            nRoots = nRoots + 1
        End If
    Next v
    ' plus every %USERPROFILE%\OneDrive* folder (covers unusual setups)
    Dim profile As String, d As String
    profile = Environ$("USERPROFILE")
    d = Dir(profile & "\OneDrive*", vbDirectory)
    Do While d <> "" And nRoots <= 9
        roots(nRoots) = profile & "\" & d
        nRoots = nRoots + 1
        d = Dir()
    Loop

    ' drop the protocol and host, then try progressively shorter URL tails
    ' against each root; the first (longest) tail that exists is the mapping
    Dim url As String, parts() As String, i As Long, j As Long, k As Long
    url = Replace(p, "%20", " ")
    i = InStr(url, "//")
    parts = Split(Mid$(url, i + 2), "/")     ' parts(0) = host
    For i = 1 To UBound(parts)
        Dim tail As String
        tail = ""
        For j = i To UBound(parts)
            tail = tail & "\" & parts(j)
        Next j
        For k = 0 To nRoots - 1
            If Dir(roots(k) & tail, vbDirectory) <> "" Then
                MapUrlToLocal = roots(k) & tail
                Exit Function
            End If
        Next k
    Next i
    MapUrlToLocal = ""                        ' no local match found
End Function

Public Sub ImportAll()
    Dim basePath As String
    basePath = WorkbookLocalPath()
    If basePath = "" Then
        If Not SilentMode Then
            MsgBox "The local folder of this workbook was not found." & vbCrLf & _
                   "Excel path: " & ThisWorkbook.Path & vbCrLf & vbCrLf & _
                   "Save the workbook to disk, or copy the audit folder to a " & _
                   "local disk (for example C:\Audit) and open it from there.", _
                   vbExclamation, "M365 Audit"
        End If
        Exit Sub
    End If
    ImportAllFrom basePath & "\DATA", basePath & "\MANUAL"
End Sub

' Entry point for automation: same import, no message boxes.
Public Sub ImportAllSilent()
    SilentMode = True
    ImportAll
    SilentMode = False
End Sub

Public Sub ImportAllFrom(ByVal dataPath As String, ByVal manualPath As String)
    Dim okCount As Long, failCount As Long, missCount As Long, totalRows As Long
    Dim logRow As Long
    Dim screenState As Boolean

    screenState = Application.ScreenUpdating
    On Error GoTo CleanExit
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' collector per-dataset status (optional file)
    Dim collStatus As Object, collMessage As Object
    Set collStatus = CreateObject("Scripting.Dictionary")
    Set collMessage = CreateObject("Scripting.Dictionary")
    LoadCollectorStatus dataPath & "\datasets_status.csv", collStatus, collMessage

    ClearRunInfoTable
    logRow = 11

    Dim entry As Variant, parts() As String
    For Each entry In DatasetTable()
        parts = Split(CStr(entry), "|")
        Dim csvFile As String, sheetName As String, maxRows As Long, types As String
        csvFile = dataPath & "\" & parts(0)
        sheetName = parts(1)
        maxRows = CLng(parts(2))
        types = parts(3)

        Dim status As String, message As String, rowsImported As Long
        rowsImported = 0
        If Dir(csvFile) = "" Then
            status = "MISSING": message = parts(0) & " not found in DATA folder"
            missCount = missCount + 1
        Else
            ImportOneCsv csvFile, sheetName, Len(types), maxRows, types, True, rowsImported, status, message
            If status = "OK" Then okCount = okCount + 1 Else failCount = failCount + 1
            totalRows = totalRows + rowsImported
        End If

        ' merge collector-side status (a dataset can be written empty after a FAIL)
        Dim dsKey As String
        dsKey = Replace(parts(0), ".csv", "")
        If collStatus.Exists(dsKey) Then
            If collStatus(dsKey) <> "OK" And status = "OK" Then
                status = collStatus(dsKey)
                message = collMessage(dsKey)
            End If
        End If

        WriteRunInfoRow logRow, dsKey, rowsImported, status, message
        logRow = logRow + 1
    Next entry

    ' manual landing sheets (optional)
    For Each entry In ManualTable()
        parts = Split(CStr(entry), "|")
        Dim mCsv As String
        mCsv = manualPath & "\" & parts(0)
        Dim mStatus As String, mMessage As String, mRows As Long
        mRows = 0
        If Dir(mCsv) = "" Then
            mStatus = "": mMessage = ""   ' optional: not reported when absent
        Else
            ImportOneCsv mCsv, parts(1), 0, CLng(parts(2)), "", (parts(3) = "1"), mRows, mStatus, mMessage
            WriteRunInfoRow logRow, "MANUAL " & Replace(parts(0), ".csv", ""), mRows, mStatus, mMessage
            logRow = logRow + 1
        End If
    Next entry

    StampRunInfo dataPath & "\run.json"

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    Dim warnCount As String
    warnCount = CStr(ThisWorkbook.Worksheets("RunInfo").Range("B8").Value)
    If Not SilentMode Then
        MsgBox "Import finished." & vbCrLf & vbCrLf & _
               "Datasets OK: " & okCount & vbCrLf & _
               "Datasets failed: " & failCount & vbCrLf & _
               "Datasets missing: " & missCount & vbCrLf & _
               "Rows imported: " & totalRows & vbCrLf & _
               "Collector warnings: " & warnCount & vbCrLf & vbCrLf & _
               "Details on the RunInfo sheet.", _
               IIf(failCount + missCount > 0, vbExclamation, vbInformation), "M365 Audit"
    End If

CleanExit:
    If Err.Number <> 0 Then
        If Not SilentMode Then MsgBox "Import error: " & Err.Description, vbCritical, "M365 Audit"
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = screenState
End Sub

' ---------------------------------------------------------------------------
' Import one CSV into one sheet. nCols = 0 means "use the sheet's header width".
' ---------------------------------------------------------------------------
Private Sub ImportOneCsv(ByVal csvFile As String, ByVal sheetName As String, _
                         ByVal nCols As Long, ByVal maxRows As Long, ByVal types As String, _
                         ByVal assertHeaders As Boolean, _
                         ByRef rowsImported As Long, ByRef status As String, ByRef message As String)
    Dim csvWb As Workbook, csvWs As Worksheet, destWs As Worksheet
    Dim c As Long, nRows As Long
    status = "OK": message = ""

    On Error GoTo Fail
    Set destWs = ThisWorkbook.Worksheets(sheetName)
    If nCols = 0 Then
        nCols = destWs.Cells(1, destWs.Columns.Count).End(xlToLeft).Column
        ' manual sheets carry an instruction cell 2 columns right of the headers
        Do While nCols > 1 And InStr(1, CStr(destWs.Cells(1, nCols).Value), "MANUAL INPUT") > 0
            nCols = nCols - 1
        Loop
        Do While nCols > 1 And Trim$(CStr(destWs.Cells(1, nCols).Value)) = ""
            nCols = nCols - 1
        Loop
    End If

    ' field types: explicit per column so regional settings cannot corrupt data
    Dim fi() As Variant
    ReDim fi(0 To nCols - 1)
    For c = 1 To nCols
        Dim code As Long
        code = 2 ' default: text
        If Len(types) >= c Then
            Select Case Mid$(types, c, 1)
                Case "D": code = 5      ' xlYMDFormat (ISO dates)
                Case "N", "B": code = 1 ' general (numbers, TRUE/FALSE booleans)
                Case Else: code = 2     ' text
            End Select
        Else
            code = 1                    ' manual sheets: general
        End If
        fi(c - 1) = Array(c, code)
    Next c

    Workbooks.OpenText fileName:=csvFile, Origin:=65001, StartRow:=1, _
        DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=True, _
        Space:=False, Other:=False, FieldInfo:=fi, Local:=False
    Set csvWb = ActiveWorkbook
    Set csvWs = csvWb.Worksheets(1)

    ' header check: importing into the wrong shape must fail loudly, not silently
    If assertHeaders Then
        For c = 1 To nCols
            If LCase$(Trim$(CStr(csvWs.Cells(1, c).Value))) <> LCase$(Trim$(CStr(destWs.Cells(1, c).Value))) Then
                status = "FAIL"
                message = "Header mismatch in " & Mid$(csvFile, InStrRev(csvFile, "\") + 1) & _
                          " column " & c & ": got '" & CStr(csvWs.Cells(1, c).Value) & _
                          "', expected '" & CStr(destWs.Cells(1, c).Value) & _
                          "'. File not imported - collector and workbook versions must match."
                GoTo CloseCsv
            End If
        Next c
    End If

    nRows = csvWs.UsedRange.Rows.Count - 1   ' minus header
    If nRows > maxRows Then
        message = "File has " & nRows & " rows; only the first " & maxRows & " were imported."
        status = "WARN"
        nRows = maxRows
    End If

    ' clear ONLY the contract columns: helper formula columns must survive
    destWs.Range(destWs.Cells(2, 1), destWs.Cells(maxRows + 1, nCols)).ClearContents
    If nRows > 0 Then
        destWs.Range("A2").Resize(nRows, nCols).Value = csvWs.Range("A2").Resize(nRows, nCols).Value
    End If
    rowsImported = nRows

CloseCsv:
    csvWb.Close SaveChanges:=False
    Exit Sub

Fail:
    status = "FAIL"
    message = Err.Description
    On Error Resume Next
    If Not csvWb Is Nothing Then csvWb.Close SaveChanges:=False
End Sub

' ---------------------------------------------------------------------------
' RunInfo sheet
' ---------------------------------------------------------------------------
Private Sub ClearRunInfoTable()
    ThisWorkbook.Worksheets("RunInfo").Range("A11:D" & CStr(11 + MAX_LOG_ROWS)).ClearContents
End Sub

Private Sub WriteRunInfoRow(ByVal r As Long, ByVal name As String, ByVal rows As Long, _
                            ByVal status As String, ByVal message As String)
    With ThisWorkbook.Worksheets("RunInfo")
        .Cells(r, 1).Value = name
        .Cells(r, 2).Value = rows
        .Cells(r, 3).Value = status
        .Cells(r, 4).Value = message
    End With
End Sub

Private Sub StampRunInfo(ByVal runJsonPath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("RunInfo")
    If Dir(runJsonPath) = "" Then
        ws.Range("B8").Value = "run.json not found"
        Exit Sub
    End If
    Dim txt As String
    txt = ReadUtf8File(runJsonPath)
    ws.Range("B3").Value = JsonValue(txt, "tenantDisplayName")
    ws.Range("B5").Value = JsonValue(txt, "account")
    ws.Range("B6").Value = JsonValue(txt, "tenantId")
    ws.Range("B7").Value = IIf(LCase$(JsonValue(txt, "concealedNames")) = "true", "Yes - per-user reports are pseudonymized", "No")
    ws.Range("B8").Value = Val(JsonValue(txt, "warningsCount"))

    Dim iso As String
    iso = JsonValue(txt, "runStartUtc")
    If Len(iso) >= 10 Then
        ws.Range("B4").Value = DateSerial(CInt(Mid$(iso, 1, 4)), CInt(Mid$(iso, 6, 2)), CInt(Mid$(iso, 9, 2)))
    End If
End Sub

Private Sub LoadCollectorStatus(ByVal path As String, ByRef statusDict As Object, ByRef msgDict As Object)
    If Dir(path) = "" Then Exit Sub
    On Error Resume Next
    Dim txt As String, lines() As String, i As Long, fields() As String
    txt = ReadUtf8File(path)
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    For i = 1 To UBound(lines)   ' skip header
        If Trim$(lines(i)) <> "" Then
            fields = SplitCsvLine(lines(i))
            If UBound(fields) >= 3 Then
                statusDict(fields(0)) = fields(2)
                msgDict(fields(0)) = fields(3)
            End If
        End If
    Next i
End Sub

' ---------------------------------------------------------------------------
' Small utilities (no external references)
' ---------------------------------------------------------------------------
Private Function ReadUtf8File(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2            ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadUtf8File = stream.ReadText(-1)
    stream.Close
End Function

' Extracts "key": <string|number|bool> from a JSON text (flat keys only).
Private Function JsonValue(ByVal json As String, ByVal key As String) As String
    Dim p As Long, q As Long, ch As String
    p = InStr(1, json, """" & key & """")
    If p = 0 Then Exit Function
    p = InStr(p, json, ":")
    If p = 0 Then Exit Function
    p = p + 1
    Do While p <= Len(json) And (Mid$(json, p, 1) = " " Or Mid$(json, p, 1) = vbTab Or _
          Mid$(json, p, 1) = vbCr Or Mid$(json, p, 1) = vbLf)
        p = p + 1
    Loop
    ch = Mid$(json, p, 1)
    If ch = """" Then
        q = p + 1
        Do While q <= Len(json)
            If Mid$(json, q, 1) = """" And Mid$(json, q - 1, 1) <> "\" Then Exit Do
            q = q + 1
        Loop
        JsonValue = Replace(Mid$(json, p + 1, q - p - 1), "\""", """")
    Else
        q = p
        Do While q <= Len(json) And InStr(1, ",}]" & vbCr & vbLf, Mid$(json, q, 1)) = 0
            q = q + 1
        Loop
        JsonValue = Trim$(Mid$(json, p, q - p))
    End If
End Function

' Splits one CSV line where every field is double-quoted (collector format).
Private Function SplitCsvLine(ByVal line As String) As String()
    Dim out() As String, n As Long, i As Long, inQ As Boolean, cur As String
    ReDim out(0 To 63)
    n = 0: cur = "": inQ = False
    For i = 1 To Len(line)
        Dim ch As String
        ch = Mid$(line, i, 1)
        If inQ Then
            If ch = """" Then
                If i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                    cur = cur & """": i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQ = True
            ElseIf ch = "," Then
                out(n) = cur: n = n + 1: cur = ""
            Else
                cur = cur & ch
            End If
        End If
    Next i
    out(n) = cur
    ReDim Preserve out(0 To n)
    SplitCsvLine = out
End Function
