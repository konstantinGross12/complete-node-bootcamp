Sub CreateBasicSlidesWithSections()
    Dim pptLayout As CustomLayout
    Dim slideTitle As String
    Dim slideSubtitle As String
    Dim i As Integer
    Dim slide As slide
    
    ' Define the slide titles and subtitles for the 6 sections
    Dim titles(1 To 6) As String
    Dim subtitles(1 To 6) As String
    
    titles(1) = "Introduction"
    subtitles(1) = "Overview of the Presentation"
    
    titles(2) = "Background"
    subtitles(2) = "Context and History"
    
    titles(3) = "Current Situation"
    subtitles(3) = "Where We Stand Today"
    
    titles(4) = "Analysis"
    subtitles(4) = "Insights and Challenges"
    
    titles(5) = "Solution"
    subtitles(5) = "Proposed Actions and Strategies"
    
    titles(6) = "Conclusion"
    subtitles(6) = "Summary and Next Steps"
    
    ' Set the layout for the slides
    Set pptLayout = ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1)
    
    ' Create the slides
    For i = 1 To 6
        Set slide = ActivePresentation.Slides.AddSlide(ActivePresentation.Slides.Count + 1, pptLayout)
        slide.Shapes(1).TextFrame.TextRange.Text = titles(i)
        slide.Shapes(2).TextFrame.TextRange.Text = subtitles(i)
        
        ' Optional: Add a header to each slide
        slide.Shapes.Title.TextFrame.TextRange.Text = "Section " & i & ": " & titles(i)
    Next i
    
    MsgBox "6 Basic Slides Created Successfully!"
End Sub