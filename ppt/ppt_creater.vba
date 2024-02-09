Sub CreateSlide(ByVal pres As Object, ByVal title As String, ByVal subtitle As String, ByVal slideIndex As Integer)
    ' Add a new slide (assuming Title and Subtitle layout)
    Dim slide As Object
    Set slide = pres.Slides.Add(slideIndex, 1)
    
    ' Set title and subtitle for each slide
    slide.Shapes(1).TextFrame.TextRange.Text = title
    slide.Shapes(2).TextFrame.TextRange.Text = subtitle
End Sub

Sub CreatePresentation()
    Dim ppt As Object
    Dim pres As Object
    Dim slideIndex As Integer

    ' Create a new PowerPoint application and presentation
    Set ppt = CreateObject("PowerPoint.Application")
    Set pres = ppt.Presentations.Add
    ppt.Visible = True

    ' Initialize slideIndex
    slideIndex = 1

    ' Slide 1
    CreateSlide pres, "Slide 1", "Slide 1 Subtitle", slideIndex
    slideIndex = slideIndex + 1

    ' Slide 2
    CreateSlide pres, "Slide 2", "Slide 2 Subtitle", slideIndex
    slideIndex = slideIndex + 1

    ' Slide 3
    CreateSlide pres, "Slide 3", "Slide 3 Subtitle", slideIndex
    slideIndex = slideIndex + 1

    ' Slide 4
    CreateSlide pres, "Slide 4", "Slide 4 Subtitle", slideIndex
    slideIndex = slideIndex + 1

    ' Slide 5
    CreateSlide pres, "Slide 5", "Slide 5 Subtitle", slideIndex
    slideIndex = slideIndex + 1

    ' Slide 6
    CreateSlide pres, "Slide 6", "Slide 6 Subtitle", slideIndex
    slideIndex = slideIndex + 1

    ' Slide 7
    CreateSlide pres, "Slide 7", "Slide 7 Subtitle", slideIndex
    slideIndex = slideIndex + 1

    ' Slide 8
    CreateSlide pres, "Slide 8", "Slide 8 Subtitle", slideIndex
    slideIndex = slideIndex + 1

    ' Slide 9
    CreateSlide pres, "Slide 9", "Slide 9 Subtitle", slideIndex
    slideIndex = slideIndex + 1

    ' Slide 10
    CreateSlide pres, "Slide 10", "Slide 10 Subtitle", slideIndex
    slideIndex = slideIndex + 1

End Sub
