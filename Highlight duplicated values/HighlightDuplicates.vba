Sub HighlightDuplicates()
    'Declare your variables
        Dim MyRange As Range
        Dim MyCell As Range
    'Define the target Range.
        Set MyRange = Selection 
    'Start looping through the range.
        For Each MyCell In MyRange 
    'Ensure the cell has Text formatting.
            If WorksheetFunction.CountIf(MyRange, MyCell.Value) > 1 Then
                MyCell.Interior.ColorIndex = 36
            End If
    'Get the next cell in the range
        Next MyCell
End Sub

