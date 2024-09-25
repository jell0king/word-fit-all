Attribute VB_Name = "Fit_all"
Sub FitallImagesOnOnePage()
    Dim wdDoc As Document
    Dim wdPageWidth As Single
    Dim wdPageHeight As Single
    Dim scaleFactor As Single
    Dim selectedShape As InlineShape
    Dim totalWidth As Single
    Dim totalHeight As Single
    Dim count As Integer
    Dim currentLeft As Single
    Dim currentTop As Single
    
    Set wdDoc = ActiveDocument
    
    Selection.WholeStory
    
    wdPageWidth = wdDoc.Sections(1).PageSetup.PageWidth * 0.78
    wdPageHeight = wdDoc.Sections(1).PageSetup.PageHeight * 0.76
    
    totalWidth = 0
    totalHeight = 0
    count = 0
    
    For Each selectedShape In Selection.InlineShapes
        If selectedShape.Type = wdInlineShapePicture Then
            totalWidth = totalWidth + selectedShape.Width
            totalHeight = totalHeight + selectedShape.Height
            count = count + 1
        End If
    Next selectedShape
    
    If count > 0 Then
    
        countOnLine = Sqr(count)
        If Int(countOnLine) <> countOnLine Then
            countOnLine = Int(countOnLine) + 1
        End If
        totalWidth = totalWidth / count * countOnLine
        totalHeight = totalWidth / count * countOnLine
        
        scaleFactorWidth = wdPageWidth / totalWidth
        scaleFactorHeight = wdPageHeight / totalHeight
        If scaleFactorWidth > scaleFactorHeight Then
            scaleFactor = scaleFactorHeight
        Else
            scaleFactor = scaleFactorWidth
        End If
        
        For Each selectedShape In Selection.InlineShapes
            If selectedShape.Type = wdInlineShapePicture Then
                selectedShape.LockAspectRatio = msoFalse
                selectedShape.Width = selectedShape.Width * scaleFactor
                selectedShape.Height = selectedShape.Height * scaleFactor
            End If
        Next selectedShape
    End If
End Sub

