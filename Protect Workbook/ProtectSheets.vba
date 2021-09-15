Sub ProtectSheets()
    'Declare your variables
        Dim ws As Worksheet
    'Start looping through all worksheets
        For Each ws In ActiveWorkbook.Worksheets
    'Protect and loop to next worksheet
        ws.Protect Password:="1234"
        Next ws
End Sub