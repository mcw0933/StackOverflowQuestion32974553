VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents MyQueryTable As QueryTable
Attribute MyQueryTable.VB_VarHelpID = -1

Private Sub MyQueryTable_AfterRefresh(ByVal Success As Boolean)
    If Success Then
        MsgBox "done!"
    Else
        MsgBox "kaboom!"
    End If
End Sub

Friend Sub HookUpQueryTable(qt As QueryTable)
    Set MyQueryTable = qt
End Sub

