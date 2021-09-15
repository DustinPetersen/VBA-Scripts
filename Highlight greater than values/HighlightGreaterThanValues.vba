Sub HighlightGreaterThanValues()
    Dim i As Integer
    i = InputBox("Enter Greater Than Value", "Enter Value")
    Selection.FormatConditions.Delete
    'Change the Operator to xlLower to highlight lower than values
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=i
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .Font.Color = RGB(, , )
            .Interior.Color = RGB(31, 218, 154)
        End With
End Sub
