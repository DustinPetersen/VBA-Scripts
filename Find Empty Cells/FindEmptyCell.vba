Sub FindEmptyCell()
    ActiveCell.Offset(1, ).Select
       Do While Not IsEmpty(ActiveCell)
          ActiveCell.Offset(1, ).Select
       Loop
End Sub
