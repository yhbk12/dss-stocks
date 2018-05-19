Attribute VB_Name = "Module1"
Sub RetrieveHistoricalData()
Attribute RetrieveHistoricalData.VB_ProcData.VB_Invoke_Func = " \n14"

    Columns("D:J").EntireColumn.ClearContents

    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://www.google.com/finance/historical?q=NASDAQ:" & ActiveSheet.Range("B1").Value & "&output=csv" _
        , Destination:=Range("D1"))
        .PostText = "GoogleFinanceHistorical_2"
        .Name = "GoogleFinanceQuery"
        .FieldNames = False
        .RefreshStyle = xlInsertDeleteCells
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .HasAutoFormat = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .TablesOnlyFromHTML = True
        .SaveData = True
        .Refresh BackgroundQuery:=False
        .UseListObject = False
    End With
    
    DeleteWebQueries

    Columns("D:D").Select
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 4), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1))
    ActiveSheet.Range("D1") = "Date"
    Columns("D").EntireColumn.AutoFit

End Sub

Sub DeleteWebQueries()
    Dim i As Integer
    With ActiveSheet
        For i = .QueryTables.Count To 1 Step -1
            .QueryTables(i).Delete
        Next
    End With
End Sub
