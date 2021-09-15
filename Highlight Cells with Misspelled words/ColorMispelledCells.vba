 Sub ColorMispelledCells()
    For Each cl In ActiveSheet.UsedRange
        If Not Application.CheckSpelling(Word:=cl.Text) Then _
        cl.Interior.ColorIndex = 28
    Next cl
End Sub