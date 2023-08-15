Sub TextboxesToShapes()
    Dim textBox As Shape
    Dim slide As slide
    Dim rectangleCut As Shape
    Dim mergeShapes(0 To 1) As Integer
    Dim mergeRange As ShapeRange
    Dim tbArray() As Integer
    
    ' Get the active slide
    Set slide = ActiveWindow.View.slide
    
    
    If ActiveWindow.Selection.Type = ppSelectionText Or ActiveWindow.Selection.Type = ppSelectionShapes Then
        
        ' Save the list of selected shapes in a separate array with shape indexes
        cntSelected = ActiveWindow.Selection.ShapeRange.Count
        ReDim tbArray(cntSelected)
        For i = 1 To cntSelected
            For j = 1 To slide.Shapes.Count
                If ActiveWindow.Selection.ShapeRange(i).Id = slide.Shapes(j).Id Then
                    tbArray(i - 1) = j
                End If
            Next j
        Next i
    
        ' Loop through each selected textbox
        For i = 0 To cntSelected - 1
            
            'Get the index of the current textbox and find the textbox itself
            currIndex = tbArray(i)
            Set textBox = slide.Shapes(currIndex)
           
            ' Get dimensions of the current textbox
            currLeft = textBox.Left
            currTop = textBox.Top
            currWidth = textBox.Width
            currHeight = textBox.Height
            
            ' Create a rectangle with exact same dimensions
            Set rectangleCut = slide.Shapes.AddShape(Type:=msoShapeRectangle, Left:=currLeft, Top:=currTop, Width:=currWidth, Height:=currHeight)
            
            ' Find the index of the newly created rectangle
            rectIndex = 0
            For k = 1 To slide.Shapes.Count
                If rectangleCut.Id = slide.Shapes(k).Id Then
                    rectIndex = k
                End If
            Next k
            
            ' Create a ShapeRange with the current textbox and the rectangle
            mergeShapes(0) = currIndex
            mergeShapes(1) = rectIndex
            Set mergeRange = slide.Shapes.Range(mergeShapes)
            
            ' Intersect the shapes with the textbox as a primary shape
            mergeRange.mergeShapes msoMergeIntersect, textBox
            
        Next i
    Else
        MsgBox "Please select a shape."
        Exit Sub
    End If
    
End Sub
