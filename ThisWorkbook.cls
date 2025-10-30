VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ApplyPermissionsAll
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Sh.Name = SHEET_NAME Then ApplyPermissionsAll
End Sub
