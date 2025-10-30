Attribute VB_Name = "Module1"
Option Explicit

Public Const PWD As String = "1234"
Public Const START_ROW As Long = 7
Public Const TARGET_SHEET As String = "For Lathe Tooling (USEd)"

Private Function GetTargetSheet() As Worksheet
    On Error Resume Next
    Set GetTargetSheet = ThisWorkbook.Worksheets(TARGET_SHEET)
    On Error GoTo 0
End Function

Public Sub ApplyPermissionsAll(Optional ByVal ws As Worksheet = Nothing)
    Dim r As Long, v As String, kVal As String, lr As Long

    If ws Is Nothing Then Set ws = GetTargetSheet()
    If ws Is Nothing Then
        MsgBox "Sheet '" & TARGET_SHEET & "' not found.", vbCritical
        Exit Sub
    End If
    If ws.Name <> TARGET_SHEET Then Exit Sub

    If ws.ProtectContents Then
        On Error Resume Next
        ws.Unprotect Password:=PWD
        If ws.ProtectContents Then ws.Unprotect
        On Error GoTo 0
        If ws.ProtectContents Then
            MsgBox "Cannot unprotect '" & ws.Name & _
                   "'. Password may be different.", vbCritical
            Exit Sub
        End If
    End If

    ws.Cells.Locked = True

    lr = LastUsedRow(ws, START_ROW)
    ws.Range("C" & START_ROW & ":C" & lr).Locked = False
    ws.Range("K" & START_ROW & ":K" & lr).Locked = False
    ws.Range("L" & START_ROW & ":L" & lr).Locked = False
    ws.Range("M" & START_ROW & ":M" & lr).Locked = False

    Dim rEnd As Long: rEnd = lr
    For r = START_ROW To rEnd
        v = UCase$(Trim$(ws.Cells(r, "C").Value))
        kVal = UCase$(Trim$(ws.Cells(r, "K").Value))

        If v = "NG" Or v = "RP" Then
            ws.Range("B" & r & ":M" & r).Locked = True
            ws.Cells(r, "C").Locked = False
        Else
            ws.Cells(r, "K").Locked = False
            If kVal = "YES" Then
                ws.Range("L" & r & ":M" & r).Locked = True
            Else
                ws.Cells(r, "L").Locked = False
                ws.Cells(r, "M").Locked = False
            End If
        End If
    Next r

    ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Private Function LastUsedRow(ws As Worksheet, startRow As Long) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find("*", ws.[A1], xlFormulas, , xlByRows, xlPrevious)
    On Error GoTo 0
    LastUsedRow = IIf(f Is Nothing, startRow, Application.Max(f.Row, startRow))
End Function

Public Sub UnlockSheet(Optional ByVal ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = GetTargetSheet()
    If ws Is Nothing Or ws.Name <> TARGET_SHEET Then Exit Sub

    Dim p As String: p = InputBox("Enter password to unprotect the sheet:", "Unlock")
    If p = vbNullString Then Exit Sub
    On Error GoTo Wrong
    ws.Unprotect Password:=p
    MsgBox "Sheet is now unprotected.", vbInformation
    Exit Sub
Wrong:
    MsgBox "Wrong password.", vbCritical
End Sub

Public Sub ProtectSheet(Optional ByVal ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = GetTargetSheet()
    If ws Is Nothing Or ws.Name <> TARGET_SHEET Then Exit Sub

    ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
    MsgBox "Sheet protected.", vbInformation
End Sub
