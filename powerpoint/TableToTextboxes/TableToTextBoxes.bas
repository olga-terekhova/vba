Sub TableToTextboxes()
    Dim tbl As Table
    Dim cell As cell
    Dim textBox As Shape
    Dim slide As slide
    Dim numCol
    Dim numRows
    
    ' Get the active slide
    Set slide = ActiveWindow.View.slide
    
    ' Check if one table is selected
    If ActiveWindow.Selection.Type = ppSelectionText Or ActiveWindow.Selection.Type = ppSelectionShapes Then
        If ActiveWindow.Selection.ShapeRange.Count = 1 Then
            If ActiveWindow.Selection.ShapeRange(1).HasTable Then
                Set tableShape = ActiveWindow.Selection.ShapeRange(1)
            Else
                MsgBox "Please select a single table."
                Exit Sub
            End If
        Else
            MsgBox "Please select only one table."
            Exit Sub
        End If
    Else
        MsgBox "Please select a single table."
        Exit Sub
    End If
   
    
    ' Set table reference
    Set tbl = tableShape.Table
    numCol = tbl.Columns.Count
    numRows = tbl.Rows.Count
    
    ' Loop through all cells
    For i = 1 To numCol
        For j = 1 To numRows
            
            ' Record current cell properties
            Set cell = tbl.cell(i, j)
            cellWidth = cell.Shape.Width
            cellHeight = cell.Shape.Height
            cellLeft = cell.Shape.Left
            cellTop = cell.Shape.Top
            cellFontName = cell.Shape.TextFrame2.TextRange.Font.Name
            cellFontSize = cell.Shape.TextFrame2.TextRange.Font.Size
            cellFontSpacing = cell.Shape.TextFrame2.TextRange.Font.Spacing
            cellFontColor = cell.Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
            cellAlignment = cell.Shape.TextFrame2.TextRange.ParagraphFormat.Alignment
            cellVerticalAnchor = cell.Shape.TextFrame2.VerticalAnchor
            cellMarginLeft = cell.Shape.TextFrame2.MarginLeft
            cellMarginRight = cell.Shape.TextFrame2.MarginRight
            cellMarginTop = cell.Shape.TextFrame2.MarginTop
            cellMarginBottom = cell.Shape.TextFrame2.MarginBottom
            cellText = cell.Shape.TextFrame2.TextRange.Text
            cellForeColor = cell.Shape.Fill.ForeColor.RGB
            
            'Add a text box
            Set textBox = slide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                              Left:=cellLeft, Top:=cellTop, _
                                              Width:=cellWidth, Height:=cellHeight)
            
            'Assign properties from the current cell
            textBox.TextFrame2.AutoSize = msoAutoSizeNone
            textBox.TextFrame2.TextRange.Font.Name = cellFontName
            textBox.TextFrame2.TextRange.Font.Size = cellFontSize
            textBox.TextFrame2.TextRange.Font.Spacing = cellFontSpacing
            textBox.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = cellFontColor
            textBox.TextFrame2.MarginLeft = cellMarginLeft
            textBox.TextFrame2.MarginRight = cellMarginRight
            textBox.TextFrame2.MarginTop = cellMarginTop
            textBox.TextFrame2.MarginBottom = cellMarginBottom
            textBox.TextFrame2.TextRange.Text = cellText
            textBox.TextFrame2.TextRange.ParagraphFormat.Alignment = cellAlignment
            textBox.TextFrame2.VerticalAnchor = cellVerticalAnchor
            textBox.Fill.ForeColor.RGB = cellForeColor
            
        Next j
    Next i
    
End Sub
