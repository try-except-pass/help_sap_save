Attribute VB_Name = "modWatcher"
Option Explicit

'==============================================================================
' Save As watcher. Runs inside a hidden Excel started by modExport's StartWatcher.
' Finds the native "Save As" dialog and fills it without needing focus.
'
' Lives in Watcher.xlsm. Keep this module named "modWatcher": the launcher
' calls modWatcher.WatchSaveAs.
'==============================================================================

'--- Windows API ---------------------------------------------------------------
Private Declare PtrSafe Function FindWindowExW Lib "user32" (ByVal hParent As LongPtr, ByVal hChildAfter As LongPtr, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClassNameW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetDlgCtrlID Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetDlgItem Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function PostMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const WM_SETTEXT As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const WM_COMMAND As Long = &H111

'--- Dialog identification ----------------------------------------------------
Private Const DLG_CLASS As String = "#32770"
Private Const DLG_TITLE As String = "Save As"            ' compared case-insensitively
Private Const ID_SAVE_BUTTON As Long = 1                 ' IDOK
Private Const ID_FILENAME_EDIT_NEW As Long = &H3E9       ' File name edit in the newer dialog (inside a ComboBox)
Private Const ID_FILENAME_COMBO As Long = &H47C          ' cmb13: File name combo (older dialogs)
Private Const ID_FILENAME_EDIT As Long = &H480           ' edt1: File name edit (oldest dialogs)

Private Const DEFAULT_TIMEOUT_SEC As Long = 300
Private Const POLL_MS As Long = 250

'==============================================================================
' Entry point, called by the launcher .vbs
'==============================================================================

Public Sub WatchSaveAs(ByVal targetPath As String, Optional ByVal timeoutSec As Long = DEFAULT_TIMEOUT_SEC)
    Dim i As Long

    On Error Resume Next
    WritePid GetCurrentProcessId()   ' left in place: KillWatcher in modExport deletes it

    For i = 1 To (timeoutSec * 1000) \ POLL_MS
        If TryFillDialog(targetPath) Then Exit For
        Sleep POLL_MS
    Next i
End Sub

'==============================================================================
' Dialog handling
'==============================================================================

' Finds a visible "Save As" dialog, writes the path into File name, confirms the
' text stuck, then triggers Save. Works without the dialog having focus.
Private Function TryFillDialog(ByVal targetPath As String) As Boolean
    Dim dlg As LongPtr, saveBtn As LongPtr, nameEdit As LongPtr

    dlg = FindWindowExW(0, 0, StrPtr(DLG_CLASS), 0)
    Do While dlg <> 0
        If IsWindowVisible(dlg) <> 0 And StrComp(GetText(dlg), DLG_TITLE, vbTextCompare) = 0 Then
            saveBtn = GetDlgItem(dlg, ID_SAVE_BUTTON)
            nameEdit = FindFileNameEdit(dlg)

            If saveBtn <> 0 And nameEdit <> 0 Then
                SendMessageW nameEdit, WM_SETTEXT, 0, StrPtr(targetPath)

                If GetText(nameEdit) = targetPath Then
                    PostMessageW dlg, WM_COMMAND, ID_SAVE_BUTTON, saveBtn
                    TryFillDialog = True
                    Exit Function
                End If
            End If
        End If
        dlg = FindWindowExW(0, dlg, StrPtr(DLG_CLASS), 0)
    Loop
End Function

' Searches all child windows for the File name Edit control.
' Newer dialog: Edit id 0x3E9 inside a ComboBox. Older dialogs: combo id 0x47C or edit id 0x480.
Private Function FindFileNameEdit(ByVal parent As LongPtr) As LongPtr
    Dim child As LongPtr, p As LongPtr, found As LongPtr

    child = FindWindowExW(parent, 0, 0, 0)
    Do While child <> 0
        If ClassOf(child) = "Edit" Then
            p = GetParent(child)
            If (GetDlgCtrlID(child) = ID_FILENAME_EDIT_NEW And ClassOf(p) = "ComboBox") _
               Or GetDlgCtrlID(child) = ID_FILENAME_EDIT _
               Or GetDlgCtrlID(p) = ID_FILENAME_COMBO _
               Or GetDlgCtrlID(GetParent(p)) = ID_FILENAME_COMBO Then
                FindFileNameEdit = child
                Exit Function
            End If
        End If

        found = FindFileNameEdit(child)
        If found <> 0 Then
            FindFileNameEdit = found
            Exit Function
        End If

        child = FindWindowExW(parent, child, 0, 0)
    Loop
End Function

'==============================================================================
' Helpers
'==============================================================================

Private Function ClassOf(ByVal hWnd As LongPtr) As String
    Dim buf As String, n As Long
    buf = String$(256, vbNullChar)
    n = GetClassNameW(hWnd, StrPtr(buf), 256)
    ClassOf = Left$(buf, n)
End Function

Private Function GetText(ByVal hWnd As LongPtr) As String
    Dim buf As String, n As Long
    buf = String$(1024, vbNullChar)
    n = CLng(SendMessageW(hWnd, WM_GETTEXT, 1024, StrPtr(buf)))
    GetText = Left$(buf, n)
End Function

' Must match PidPath() in modExport
Private Function PidPath() As String
    PidPath = Environ$("TEMP") & "\saveas_watcher.pid"
End Function

Private Sub WritePid(ByVal pid As Long)
    Dim f As Integer
    f = FreeFile
    Open PidPath() For Output As #f
    Print #f, pid
    Close #f
End Sub

'==============================================================================
' Diagnostics: open Watcher.xlsm, open a Save As dialog, run from Immediate
'==============================================================================

' Lists every visible dialog and whether the watcher can see its controls.
Public Sub DiagDialogs()
    Dim dlg As LongPtr
    dlg = FindWindowExW(0, 0, StrPtr(DLG_CLASS), 0)
    Do While dlg <> 0
        If IsWindowVisible(dlg) <> 0 Then
            Debug.Print "title='" & GetText(dlg) & "'  saveBtn=" & GetDlgItem(dlg, ID_SAVE_BUTTON) & _
                        "  nameEdit=" & FindFileNameEdit(dlg)
        End If
        dlg = FindWindowExW(0, dlg, StrPtr(DLG_CLASS), 0)
    Loop
    Debug.Print "--- done"
End Sub

' Writes the full control tree of every visible dialog to %TEMP%\dialog_tree.txt
Public Sub DiagTree()
    Dim dlg As LongPtr, f As Integer, outPath As String

    outPath = Environ$("TEMP") & "\dialog_tree.txt"
    f = FreeFile
    Open outPath For Output As #f

    dlg = FindWindowExW(0, 0, StrPtr(DLG_CLASS), 0)
    Do While dlg <> 0
        If IsWindowVisible(dlg) <> 0 Then
            Print #f, "DIALOG '" & GetText(dlg) & "'"
            PrintChildren f, dlg, 1
        End If
        dlg = FindWindowExW(0, dlg, StrPtr(DLG_CLASS), 0)
    Loop

    Close #f
    Debug.Print "Written to " & outPath
End Sub

Private Sub PrintChildren(ByVal f As Integer, ByVal parent As LongPtr, ByVal depth As Long)
    Dim child As LongPtr
    child = FindWindowExW(parent, 0, 0, 0)
    Do While child <> 0
        Print #f, String$(depth * 2, " ") & ClassOf(child) & "  id=" & Hex$(GetDlgCtrlID(child)) & _
                  "  '" & Left$(GetText(child), 40) & "'"
        PrintChildren f, child, depth + 1
        child = FindWindowExW(parent, child, 0, 0)
    Loop
End Sub
