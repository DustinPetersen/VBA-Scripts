Sub PivotTableForExcel2007()
    Dim SourceRange As Range
    Set SourceRange = Sheets("Sheet1").Range("A3:N86")
    ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SourceRange, _
    Version:=xlPivotTableVersion12).CreatePivotTable _
    TableDestination:="", _
    TableName:="", _
    DefaultVersion:=xlPivotTableVersion12
End Su