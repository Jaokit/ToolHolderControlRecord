Attribute VB_Name = "Module1"
Option Explicit

Public Const PWD As String = "1234"
Public Const SHEET_NAME As String = "For Lathe Tooling"

Public Sub ApplyPermissionsAll()
    Dim ws As Worksheet, r As Long, v As String, kVal As String
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    On Error Resume Next
    ws.Unprotect Password:=PWD
    On Error GoTo 0

    ws.Cells.Locked = True
    ws.Range("C5:C16,K5:K16,L5:L16,M5:M16").Locked = False

    For r = 5 To 16
        v = UCase$(Trim$(ws.Cells(r, "C").Value))
        kVal = UCase$(Trim$(ws.Cells(r, "K").Value))

        If v = "NG" Then
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

Public Sub UnlockSheet()
    Dim ws As Worksheet, p As String
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    p = InputBox("Enter password to unprotect the sheet:", "Unlock")
    If p = vbNullString Then Exit Sub
    On Error GoTo Wrong
    ws.Unprotect Password:=p
    MsgBox "Sheet is now unprotected.", vbInformation
    Exit Sub
Wrong:
    MsgBox "Wrong password.", vbCritical
End Sub

Public Sub ProtectSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
    MsgBox "Sheet protected.", vbInformation
End Sub
