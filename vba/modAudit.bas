Attribute VB_Name = "modAudit"
Option Explicit

' ===========================================================================
' modAudit - RUN AUDIT button.
' Launches Collect-M365Audit.ps1 in a visible PowerShell window (the user must
' see the browser sign-in and the progress bars), waits for DATA\_DONE.flag,
' then imports everything through modImport.ImportAll.
' ===========================================================================

Private Const WAIT_TIMEOUT_MINUTES As Long = 120

Public Sub RunAudit()
    Dim basePath As String, dataPath As String, scriptPath As String, flagPath As String

    ' local path even when the workbook is opened from OneDrive/SharePoint
    ' (ThisWorkbook.Path returns an https:// URL there, which breaks Dir/Kill)
    basePath = modImport.WorkbookLocalPath()
    If basePath = "" Then
        MsgBox "The local folder of this workbook was not found." & vbCrLf & _
               "Excel path: " & ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "If the file has never been saved, save it first. If it is open " & _
               "from OneDrive/SharePoint, make sure the folder is synced on this " & _
               "PC, or copy the audit folder to a local disk (for example C:\Audit).", _
               vbCritical, "M365 Audit"
        Exit Sub
    End If

    scriptPath = basePath & "\Collect-M365Audit.ps1"
    If Dir(scriptPath) = "" Then
        MsgBox "Collect-M365Audit.ps1 was not found next to the workbook:" & vbCrLf & _
               scriptPath & vbCrLf & vbCrLf & _
               "Copy the script into the same folder as the workbook.", vbCritical, "M365 Audit"
        Exit Sub
    End If

    dataPath = basePath & "\DATA"
    flagPath = dataPath & "\_DONE.flag"
    If Dir(dataPath, vbDirectory) = "" Then MkDir dataPath
    If Dir(flagPath) <> "" Then Kill flagPath   ' remove stale flag from a previous run

    ' optional toggles from the Config sheet
    Dim extraArgs As String
    extraArgs = ""
    If LCase$(Trim$(CStr(ThisWorkbook.Worksheets("Config").Range("B12").Value))) = "yes" Then
        extraArgs = extraArgs & " -DeepMailboxScan"
    End If
    If LCase$(Trim$(CStr(ThisWorkbook.Worksheets("Config").Range("B10").Value))) = "yes" Then
        extraArgs = extraArgs & " -IncludeExternalFileAccess"
    End If

    Dim args As String
    args = "-NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & _
           """ -OutputPath """ & dataPath & """" & extraArgs

    ' visible window on purpose: browser sign-in + progress must be visible
    Dim taskId As Double
    On Error Resume Next
    taskId = Shell("pwsh.exe " & args, vbNormalFocus)          ' PowerShell 7 preferred
    If Err.Number <> 0 Then
        Err.Clear
        taskId = Shell("powershell.exe " & args, vbNormalFocus) ' Windows PowerShell 5.1
    End If
    On Error GoTo 0
    If taskId = 0 Then
        MsgBox "Could not start PowerShell.", vbCritical, "M365 Audit"
        Exit Sub
    End If

    ' wait for the collector, cancellable with ESC
    Dim startTime As Double, elapsed As Long, canceled As Boolean
    startTime = Timer
    canceled = False
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Interrupted

    Do While Dir(flagPath) = ""
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 2)
        elapsed = CLng((Timer - startTime + 86400) Mod 86400)
        Application.StatusBar = "M365 Audit: collecting data... " & _
            Format$(elapsed \ 60, "0") & "m" & Format$(elapsed Mod 60, "00") & "s elapsed. " & _
            "Watch the PowerShell window. Press ESC to stop waiting."
        If elapsed > WAIT_TIMEOUT_MINUTES * 60 Then
            Application.StatusBar = False
            MsgBox "Timeout after " & WAIT_TIMEOUT_MINUTES & " minutes. " & _
                   "If the collector is still running, wait for it to finish, then run the macro ImportAll.", _
                   vbExclamation, "M365 Audit"
            Exit Sub
        End If
    Loop
    GoTo FlagFound

Interrupted:
    If Err.Number = 18 Then       ' user pressed ESC
        canceled = True
    Else
        Application.StatusBar = False
        MsgBox "Unexpected error while waiting: " & Err.Description, vbCritical, "M365 Audit"
        Exit Sub
    End If

FlagFound:
    On Error GoTo 0
    Application.EnableCancelKey = xlInterrupt
    Application.StatusBar = False

    If canceled Then
        MsgBox "Waiting stopped. The PowerShell collector may still be running." & vbCrLf & _
               "When it is finished, run the macro 'ImportAll' (Alt+F8) to load the data.", _
               vbInformation, "M365 Audit"
        Exit Sub
    End If

    ' flag content: "OK ..." / "PARTIAL ..." / "FAILED ..."
    Dim flagText As String
    flagText = ReadTextFile(flagPath)
    If InStr(1, flagText, "FAILED", vbTextCompare) > 0 Then
        MsgBox "The collector failed. Last log lines:" & vbCrLf & vbCrLf & LogTail(basePath), _
               vbCritical, "M365 Audit"
        Exit Sub
    End If

    modImport.ImportAll
End Sub

' ---------------------------------------------------------------------------
Private Function ReadTextFile(ByVal path As String) As String
    On Error Resume Next
    Dim f As Integer, txt As String
    f = FreeFile
    Open path For Input As #f
    txt = Input$(LOF(f), #f)
    Close #f
    ReadTextFile = txt
End Function

Private Function LogTail(ByVal basePath As String) As String
    On Error Resume Next
    Dim logDir As String, fileName As String, newest As String, newestTime As Date
    logDir = basePath & "\logs\"
    fileName = Dir(logDir & "collector_*.log")
    Do While fileName <> ""
        If FileDateTime(logDir & fileName) > newestTime Then
            newestTime = FileDateTime(logDir & fileName)
            newest = fileName
        End If
        fileName = Dir()
    Loop
    If newest = "" Then
        LogTail = "(no log file found in " & logDir & ")"
        Exit Function
    End If

    Dim txt As String, lines() As String, i As Long, fromLine As Long, out As String
    txt = ReadTextFile(logDir & newest)
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    fromLine = UBound(lines) - 25
    If fromLine < 0 Then fromLine = 0
    For i = fromLine To UBound(lines)
        If Trim$(lines(i)) <> "" Then out = out & lines(i) & vbCrLf
    Next i
    LogTail = out & vbCrLf & "Full log: " & logDir & newest
End Function
