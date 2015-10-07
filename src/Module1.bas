Attribute VB_Name = "Module1"
Sub FetchData()
    Dim s As Worksheet
    Set s = Sheets("Data")
    
    s.QueryTables(1).Refresh
        
    s.Activate
End Sub
