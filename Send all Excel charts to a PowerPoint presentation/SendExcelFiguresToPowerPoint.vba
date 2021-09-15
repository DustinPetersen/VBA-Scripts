Sub SendExcelFiguresToPowerPoint()
    'Set reference to Microsoft Powerpoint Object Library
    'Declare your variables
        Dim PP As PowerPoint.Application
        Dim PPPres As PowerPoint.Presentation
        Dim PPSlide As PowerPoint.Slide
        Dim i As Integer
    'Check for charts; exit if no charts exist
        Sheets("Slide Data").Select
            If ActiveSheet.ChartObjects.Count < 1 Then
                MsgBox "No charts existing the active sheet"
                Exit Sub
            End If
    'Open PowerPoint and create new presentation
        Set PP = New PowerPoint.Application
        Set PPPres = PP.Presentations.Add
        PP.Visible = True
    'Start the loop based on chart count
            For i = 1 To ActiveSheet.ChartObjects.Count
            'Copy the chart as a picture
                ActiveSheet.ChartObjects(i).Chart.CopyPicture _
                Size:=xlScreen, Format:=xlPicture
                Application.Wait (Now + TimeValue("0:00:1"))
            'Count slides and add new slide as next available slide number
                ppSlideCount = PPPres.Slides.Count
                Set PPSlide = PPPres.Slides.Add(SlideCount + 1, ppLayoutBlank)
                PPSlide.Select
            'Paste the picture and adjust its position; Go to next chart
                PPSlide.Shapes.Paste.Select
                PP.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                PP.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True
            Next i
    'Memory Cleanup
        Set PPSlide = Nothing
        Set PPPres = Nothing
        Set PP = Nothing
End Sub