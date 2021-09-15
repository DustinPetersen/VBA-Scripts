Sub TopTen()
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .TopBottom = xlTop10Top
            'Change the rank here to highlight a different number of values
            .Rank = 10
            .Percent = False
        End With
        With Selection.FormatConditions(1).Font
            .Color = -16752384
            .TintAndShade = 
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13561798
            .TintAndShade = 
        End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
