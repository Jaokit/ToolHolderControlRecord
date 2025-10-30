Attribute VB_Name = "Module1"
Option Explicit

Public Const PWD As String = "1234"
Public Const START_ROW As Long = 7

Public Sub ApplyPermissionsAll(Optional ByVal ws As Worksheet = Nothing)
    Dim r As Long, v As String, kVal As String, lr As Long
    If ws Is Nothing Then Set ws = ActiveSheet

    On Error Resume Next
    ws.Unprotect Password:=PWD
    On Error GoTo 0

    ws.Cells.Locked = True
    ws.Range("C" & START_ROW & ":C" & Rows.Count).Locked = False
    ws.Range("K" & START_ROW & ":K" & Rows.Count).Locked = False
    ws.Range("L" & START_ROW & ":L" & Rows.Count).Locked = False
    ws.Range("M" & START_ROW & ":M" & Rows.Count).Locked = False

    lr = LastUsedRow(ws, START_ROW)
    For r = START_ROW To lr
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
    If ws Is Nothing Then Set ws = ActiveSheet
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
    If ws Is Nothing Then Set ws = ActiveSheet
    ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
    MsgBox "Sheet protected.", vbInformation
End Sub
