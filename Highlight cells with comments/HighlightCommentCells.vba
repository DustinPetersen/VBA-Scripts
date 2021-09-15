Sub HighlightCommentCells()
    Selection.SpecialCells(xlCellTypeComments).Select
    Selection.Style= "Note"
End Sub
