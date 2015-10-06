Attribute VB_Name = "Module1"
Sub FetchDataWithHookup()
    On Error Resume Next
    
    DoFetch True
End Sub

Sub FetchDataWithoutHookup()
    On Error Resume Next
    
    DoFetch False
End Sub

Sub DoFetch(withHookup As Boolean)
    Dim qt As QueryTable
    Set qt = Sheets("Data").QueryTables(1)
    
    If (withHookup) Then
        Dim c As Class1
        Set c = New Class1
        c.HookUpQueryTable qt
    End If
    
    qt.Refresh
        
    Sheets("Data").Activate
End Sub
