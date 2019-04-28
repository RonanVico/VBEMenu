VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Call InitVBRVTools
End Sub

Private Sub Workbook_Open()
    Call InitVBRVTools
    On Error Resume Next
    Call Application.VBE.ActiveVBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 2, 0)
    'Muda o registro do windows pra liberar acesso ao VBE ;) WE ARE HACKERS
    Call ChangeRegistry_AccessVBOM
End Sub



