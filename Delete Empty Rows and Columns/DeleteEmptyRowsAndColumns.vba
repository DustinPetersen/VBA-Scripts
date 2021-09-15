Sub DeleteEmptyRowsAndColumns()
    'Declare your variables.
        Dim MyRange As Range
        Dim iCounter As Long
    'Define the target Range.
        Set MyRange = ActiveSheet.UsedRange
        'Start reverse looping through the range of Rows.
        For iCounter = MyRange.Rows.Count To 1 Step -1
    'If entire row is empty then delete it.
           If Application.CountA(Rows(iCounter).EntireRow) =  Then
               Rows(iCounter).Delete
               'Remove comment to See which are the empty rows
               'MsgBox "row " & iCounter & " is empty"
           End If
    'Increment the counter down
        Next iCounter
    'Step 6:  Start reverse looping through the range of Columns.
        For iCounter = MyRange.Columns.Count To 1 Step -1
    'Step 7: If entire column is empty then delete it.
               If Application.CountA(Columns(iCounter).EntireColumn) =  Then
                Columns(iCounter).Delete
               End If
    'Step 8: Increment the counter down
        Next iCounter      
End Sub
