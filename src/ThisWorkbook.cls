VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim c As Class1

Private Sub Workbook_Open()
    On Error Resume Next
    
    Set c = New Class1
    c.HookUpQueryTable Sheets("Data").QueryTables(1)
End Sub
