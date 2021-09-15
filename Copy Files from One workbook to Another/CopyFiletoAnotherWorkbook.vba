Sub CopyFiletoAnotherWorkbook()
    'Copy the data
        Sheets("Example 1").Range("B4:C15").Copy
    'Create a new workbook
        Workbooks.Add
    'Paste the data
        ActiveSheet.Paste
    'Turn off application alerts
        Application.DisplayAlerts = False
    'Save the newly file. Change the name of the directory.
        ActiveWorkbook.SaveAs Filename:="C:\Temp\MyNewBook.xlsx"
    'Turn application alerts back on
        Application.DisplayAlerts = True
End Sub
