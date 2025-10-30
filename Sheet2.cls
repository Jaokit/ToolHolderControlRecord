VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const R_STATUS As String = "C5:C16"
Private Const R_STORE  As String = "K5:K16"
Private Const R_WARN   As String = "L5:M16"

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitH
    Application.EnableEvents = False

    Dim rng As Range, c As Range, r As Long

    Set rng = Intersect(Target, Me.Range(R_STATUS))
    If Not rng Is Nothing Then
        For Each c In rng.Cells
            ApplyStatusRules c.Row
            UpdateLocksForRow c.Row
        Next c
    End If

    Set rng = Intersect(Target, Me.Range(R_STORE))
    If Not rng Is Nothing Then
        For Each c In rng.Cells
            ApplyStoreRules c.Row, True
            UpdateLocksForRow c.Row
        Next c
    End If

    Set rng = Intersect(Target, Me.Range(R_WARN))
    If Not rng Is Nothing Then
        For Each c In rng.Cells
            UpdateLMCellHighlight c.Row, c.Column
            UpdateLocksForRow c.Row
        Next c
    End If

ExitH:
    Application.EnableEvents = True
End Sub

Private Sub ApplyStatusRules(ByVal r As Long)
    Dim v As String: v = UCase$(Trim$(Me.Cells(r, "C").Value))
    Dim band As Range: Set band = Me.Range("B" & r & ":M" & r)

    If v = "NG" Then
        band.Interior.Color = RGB(217, 217, 217)
        Me.Range("K" & r & ":M" & r).ClearContents
    Else
        band.Interior.ColorIndex = xlColorIndexNone
    End If
End Sub

Private Sub UpdateLocksForRow(ByVal r As Long)
    If r < 5 Or r > 16 Then Exit Sub

    Dim ws As Worksheet, v As String, kVal As String
    Set ws = Me
    v = UCase$(Trim$(ws.Cells(r, "C").Value))
    kVal = UCase$(Trim$(ws.Cells(r, "K").Value))

    ws.Unprotect Password:=PWD

    ws.Range("B" & r & ":M" & r).Locked = True
    ws.Cells(r, "C").Locked = False
    If v = "NG" Then
        ws.Range("K" & r & ":M" & r).Locked = True
    Else
        ws.Cells(r, "K").Locked = False
        If kVal = "YES" Then
            ws.Range("L" & r & ":M" & r).Locked = True
        Else
            ws.Cells(r, "L").Locked = False
            ws.Cells(r, "M").Locked = False
        End If
    End If

    ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=True, AllowFiltering:=True
End Sub

Private Sub ApplyStoreRules(ByVal r As Long, ByVal showPopup As Boolean)
    Dim kVal As String: kVal = UCase$(Trim$(Me.Cells(r, "K").Value))

    Select Case kVal
        Case "YES"
            Me.Range("L" & r & ":M" & r).ClearContents
            ClearWarnColor r
        Case "NO"
            If showPopup Then
                MsgBox "Row " & r & ": K = NO. Please fill L (M/C No.) and/or M (Issuer).", vbExclamation, "Input required"
            End If
            UpdateLMCellHighlight r, Columns("L").Column
            UpdateLMCellHighlight r, Columns("M").Column
        Case Else
            ClearWarnColor r
    End Select
End Sub

Private Sub UpdateLMCellHighlight(ByVal r As Long, ByVal colIdx As Long)
    Dim isNo As Boolean
    isNo = (UCase$(Trim$(Me.Cells(r, "K").Value)) = "NO")

    If isNo Then
        If Trim$(Me.Cells(r, colIdx).Value) = "" Then
            Me.Cells(r, colIdx).Interior.Color = vbYellow
        Else
            Me.Cells(r, colIdx).Interior.ColorIndex = xlColorIndexNone
        End If
    Else
        ClearWarnColor r
    End If
End Sub

Private Sub ClearWarnColor(ByVal r As Long)
    Me.Range("L" & r & ":M" & r).Interior.ColorIndex = xlColorIndexNone
End Sub
