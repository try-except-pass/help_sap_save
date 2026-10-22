Attribute VB_Name = "modExport"
Option Explicit

'==============================================================================
' SAP export. Starts a hidden Excel running Watcher.xlsm (modWatcher.WatchSaveAs),
' which fills the native "Save As" dialog while SAP's .Press blocks this macro.
' The exported file is then opened in THIS Excel instance.
'==============================================================================

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Leave empty to use Watcher.xlsm from the same folder as this workbook,
' or set a full path, e.g. "C:\scripts\Watcher.xlsm"
Private Const WATCHER_PATH As String = ""
Private Const WATCHER_FILE As String = "Watcher.xlsm"

Private Const DEFAULT_TIMEOUT_SEC As Long = 300

'==============================================================================
' SAP export
'==============================================================================

' Exports the report and returns it opened in this Excel instance.
Public Function ExportReport(ByVal session As Object) As Workbook
    Const TARGET_DIR As String = "C:\out\"
    Const TARGET_FILE As String = "report.xlsx"
    Dim target As String
    Dim errNum As Long, errSrc As String, errDesc As String

    target = TARGET_DIR & TARGET_FILE
    On Error GoTo Fail

    ' Start early: the hidden Excel loads while SAP navigates
    StartWatcher target

    ' ... your SAP navigation up to the export ...

    ' Blocks here while the native Save As dialog is open; the watcher fills it.
    session.findById("wnd[0]/...").Press   ' <-- your export button

    ' Sometimes SAP shows its own save dialog instead of the Windows one.
    ' Fill that one directly. Check these IDs with the SAP script recorder.
    If SapDialogOpen(session) Then
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = TARGET_DIR
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = TARGET_FILE
        session.findById("wnd[1]/tbar[0]/btn[0]").Press   ' Generate
    End If

    If Not WaitForFile(target, 60) Then
        Err.Raise vbObjectError + 1, "ExportReport", "Report was not saved: " & target
    End If

    ' Closes the hidden Excel, and any copy SAP opened in it, before we open ours
    KillWatcher

    Set ExportReport = OpenHere(target)
    Exit Function

Fail:
    ' Always close the hidden Excel, then pass the original error on
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    KillWatcher
    Err.Raise errNum, errSrc, errDesc
End Function

Private Function SapDialogOpen(ByVal session As Object) As Boolean
    Dim ctl As Object
    On Error Resume Next
    Set ctl = session.findById("wnd[1]/usr/ctxtDY_PATH", False)
    SapDialogOpen = Not ctl Is Nothing
End Function

'==============================================================================
' Opening the exported file in this instance
'==============================================================================

' Returns the file opened in THIS Excel instance. If SAP already opened it here,
' that workbook is reused. If another visible Excel still has it open (SAP's
' copy), it is opened read-only here so no "file in use" prompt appears.
Public Function OpenHere(ByVal filePath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            Set OpenHere = wb
            Exit Function
        End If
    Next wb

    If IsFileLocked(filePath) Then
        Set OpenHere = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True, Notify:=False)
    Else
        Set OpenHere = Workbooks.Open(filePath, UpdateLinks:=0)
    End If
End Function

' True if another process (e.g. another Excel) has the file open.
Private Function IsFileLocked(ByVal filePath As String) As Boolean
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open filePath For Binary Access Read Write Lock Read Write As #f
    IsFileLocked = (Err.Number <> 0)
    Close #f
End Function

'==============================================================================
' Watcher control
'==============================================================================

' Kills any previous watcher, deletes the target file (avoids the "replace?"
' popup) and starts a new watcher in a hidden Excel.
Public Sub StartWatcher(ByVal targetPath As String, Optional ByVal timeoutSec As Long = DEFAULT_TIMEOUT_SEC)
    Dim wbPath As String
    wbPath = WatcherWorkbookPath()
    If Dir(wbPath) = "" Then
        Err.Raise vbObjectError + 2, "StartWatcher", "Watcher workbook not found: " & wbPath
    End If

    KillWatcher
    If Dir(targetPath) <> "" Then Kill targetPath

    WriteLauncher
    Shell "wscript.exe """ & LauncherPath() & """ """ & targetPath & """ """ & _
          wbPath & """ " & timeoutSec, vbHide
End Sub

' Closes the hidden Excel started by the watcher and its launcher script.
' Only kills an Excel started by automation (/automation in its command line),
' so an Excel opened by a person is never touched.
Public Sub KillWatcher()
    Dim wmi As Object, p As Object
    Dim pid As Long

    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")

    pid = ReadPid()
    If pid <> 0 And pid <> GetCurrentProcessId() Then
        For Each p In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId=" & pid & _
                                    " AND Name='EXCEL.EXE'")
            If IsAutomationExcel(p) Then p.Terminate
        Next p
    End If

    For Each p In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='wscript.exe' " & _
                                "AND CommandLine LIKE '%saveas[_]watcher.vbs%'")
        p.Terminate
    Next p

    Sleep 300
    If Dir(PidPath()) <> "" Then Kill PidPath()
End Sub

' Manual cleanup: closes EVERY hidden automation Excel except this one.
' Use it to get rid of leftovers from earlier runs. Note it also closes hidden
' Excels started by other tools, if any are running.
Public Sub CleanupHiddenExcel()
    Dim wmi As Object, p As Object, n As Long

    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")

    For Each p In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='EXCEL.EXE' " & _
                                "AND ProcessId<>" & GetCurrentProcessId())
        If IsAutomationExcel(p) Then
            p.Terminate
            n = n + 1
        End If
    Next p

    Debug.Print n & " hidden Excel instance(s) closed"
End Sub

' Waits until the file exists and its size has stopped changing.
' Returns False on timeout.
Public Function WaitForFile(ByVal filePath As String, Optional ByVal timeoutSec As Long = 60) As Boolean
    Dim i As Long, lastSize As Long, curSize As Long

    lastSize = -1
    For i = 1 To timeoutSec * 2
        If Dir(filePath) <> "" Then
            curSize = FileLen(filePath)
            If curSize > 0 And curSize = lastSize Then
                WaitForFile = True
                Exit Function
            End If
            lastSize = curSize
        End If
        Sleep 500
        DoEvents
    Next i
End Function

'==============================================================================
' Helpers
'==============================================================================

' True for an Excel started by automation (CreateObject), not by a person.
Private Function IsAutomationExcel(ByVal proc As Object) As Boolean
    Dim cmd As String
    cmd = "" & proc.CommandLine
    IsAutomationExcel = InStr(1, cmd, "/automation", vbTextCompare) > 0 _
                     Or InStr(1, cmd, "-embedding", vbTextCompare) > 0
End Function

Private Function WatcherWorkbookPath() As String
    If WATCHER_PATH <> "" Then
        WatcherWorkbookPath = WATCHER_PATH
    Else
        WatcherWorkbookPath = ThisWorkbook.Path & "\" & WATCHER_FILE
    End If
End Function

Private Function LauncherPath() As String
    LauncherPath = Environ$("TEMP") & "\saveas_watcher.vbs"
End Function

' Must match PidPath() in Watcher.xlsm's modWatcher
Private Function PidPath() As String
    PidPath = Environ$("TEMP") & "\saveas_watcher.pid"
End Function

Private Function ReadPid() As Long
    Dim f As Integer, pid As Long
    On Error GoTo Done
    If Dir(PidPath()) = "" Then Exit Function
    f = FreeFile
    Open PidPath() For Input As #f
    Input #f, pid
    Close #f
    ReadPid = pid
Done:
End Function

' Generates the .vbs that starts the hidden Excel.
' Args: 0 = target path, 1 = watcher workbook, 2 = timeout in seconds
Private Sub WriteLauncher()
    Dim f As Integer
    f = FreeFile
    Open LauncherPath() For Output As #f
    Print #f, "On Error Resume Next"
    Print #f, "Dim xl, wb"
    Print #f, "Set xl = CreateObject(""Excel.Application"")"
    Print #f, "xl.Visible = False"
    Print #f, "xl.DisplayAlerts = False"
    Print #f, "xl.EnableEvents = False"
    Print #f, "Set wb = xl.Workbooks.Open(WScript.Arguments(1), 0, True)"
    Print #f, "xl.Run ""'"" & wb.Name & ""'!modWatcher.WatchSaveAs"", WScript.Arguments(0), CLng(WScript.Arguments(2))"
    Print #f, "wb.Close False"
    Print #f, "xl.Quit"
    Print #f, "Set wb = Nothing"
    Print #f, "Set xl = Nothing"
    Close #f
End Sub
