Sub TrimTheSpaces()
    'Declare your variables
        Dim MyRange As Range
        Dim MyCell As Range
    'Save the Workbook before changing cells
        Select Case MsgBox("Can't Undo this action.  " & _
                            "Save Workbook First?", vbYesNoCancel)
            Case Is = vbYes
            ThisWorkbook.Save
            Case Is = vbCancel
            Exit Sub
        End Select
    'Define the target Range.
        Set MyRange = Selection
    'Start looping through the range.
        For Each MyCell In MyRange
    'Trim the Spaces.
            If Not IsEmpty(MyCell) Then
                MyCell = Trim(MyCell)
            End If
    'Get the next cell in the range
        Next MyCell
End Sub
