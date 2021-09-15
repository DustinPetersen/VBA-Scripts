Sub ExcelTableInWord()
    'Set reference to Microsoft Word Object library
    'Declare your variables
        Dim MyRange As Excel.Range
        Dim wd As Word.Application
        Dim wdDoc As Word.Document
        Dim WdRange As Word.Range
    'Copy the defined range
       Sheets("Revenue Table").Range("B4:F10").Cop
    'Open the target Word document
        Set wd = New Word.Application
        Set wdDoc = wd.Documents.Open _
        (ThisWorkbook.Path & "\" & "PasteTable.docx")
        wd.Visible = True
    'Set focus on the target bookmark
        Set WdRange = wdDoc.Bookmarks("DataTableHere").Rangе
    'Delete the old table and paste new
        On Error Resume Next
        WdRange.Tables(1).Delete
        WdRange.Paste 'paste in the table   
    'Adjust column widths
        WdRange.Tables(1).Columns.SetWidth _
        (MyRange.Width / MyRange.Columns.Count), wdAdjustSameWidth
    'Reinsert the bookmark
        wdDoc.Bookmarks.Add "DataTableHere", WdRange
    'Memory cleanup
        Set wd = Nothing
        Set wdDoc = Nothing
        Set WdRange = Nothing
End Sub