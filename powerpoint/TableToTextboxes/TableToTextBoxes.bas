' Find the index within the slide's shapes for a given shape
Function FindShapeIndex(ByVal currSlide As slide, ByVal currShape As shape)
    For j = 1 To currSlide.Shapes.Count
                If currShape.Id = currSlide.Shapes(j).Id Then
                    FindShapeIndex = j
                    Exit Function
                End If
    Next j
    FindShapeIndex = 0
End Function

' Convert a table into a set of textboxes
Sub TableToTextboxes()
    Dim tableShape As shape
    Dim tbl As Table
    Dim cell As cell
    Dim col As Column
    Dim row As row
    Dim textBox As shape
    Dim slide As slide
    Dim numCol
    Dim numRows
    Dim linesArray() As Integer
    Dim bordersRange As ShapeRange
    Dim newBorder As shape
    
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
            cellWidth = cell.shape.Width
            cellHeight = cell.shape.Height
            cellLeft = cell.shape.Left
            cellTop = cell.shape.Top
            cellFontName = cell.shape.TextFrame2.TextRange.Font.Name
            cellFontSize = cell.shape.TextFrame2.TextRange.Font.Size
            cellFontSpacing = cell.shape.TextFrame2.TextRange.Font.Spacing
            cellFontColor = cell.shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
            cellFillVisible = cell.shape.Fill.Visible
            cellFillColor = cell.shape.Fill.ForeColor.RGB
            cellAlignment = cell.shape.TextFrame2.TextRange.ParagraphFormat.Alignment
            cellVerticalAnchor = cell.shape.TextFrame2.VerticalAnchor
            cellMarginLeft = cell.shape.TextFrame2.MarginLeft
            cellMarginRight = cell.shape.TextFrame2.MarginRight
            cellMarginTop = cell.shape.TextFrame2.MarginTop
            cellMarginBottom = cell.shape.TextFrame2.MarginBottom
            cellText = cell.shape.TextFrame2.TextRange.Text
            'cellForeColor = cell.Shape.Fill.ForeColor.RGB
            
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
            textBox.Fill.ForeColor.RGB = cellFillColor
            textBox.Fill.Visible = cellFillVisible
            textBox.TextFrame2.MarginLeft = cellMarginLeft
            textBox.TextFrame2.MarginRight = cellMarginRight
            textBox.TextFrame2.MarginTop = cellMarginTop
            textBox.TextFrame2.MarginBottom = cellMarginBottom
            textBox.TextFrame2.TextRange.Text = cellText
            textBox.TextFrame2.TextRange.ParagraphFormat.Alignment = cellAlignment
            textBox.TextFrame2.VerticalAnchor = cellVerticalAnchor
            'textBox.Fill.ForeColor.RGB = cellForeColor
            
        Next j
    Next i
    
    ' Get a style for borders from the top left cell
    borderColor = tbl.cell(1, 1).Borders(ppBorderTop).ForeColor.RGB
    borderVisible = tbl.cell(1, 1).Borders(ppBorderTop).Visible
    borderStyle = tbl.cell(1, 1).Borders(ppBorderTop).Style
    borderWeight = tbl.cell(1, 1).Borders(ppBorderTop).Weight
    
    ' Create an array to store indexes of newly created border lines
    linesNum = numCol + numRows
    ReDim linesArray(linesNum + 1) ' -1 because index starts with 0; +2 to allow for the last border on the right and on the bottom
    lineCounter = 0
    
    ' Create column borders
    colHeight = tableShape.Height
    rowWidth = tableShape.Width
    colLeft = tableShape.Left
    rowTop = tableShape.Top
    
    For i = 1 To numCol
        Set col = tbl.Columns(i)
        colWidth = col.Width
        Set newBorder = slide.Shapes.AddLine(colLeft, rowTop, colLeft, rowTop + colHeight)
        newBorder.Line.ForeColor.RGB = borderColor
        newBorder.Line.Visible = borderVisible
        newBorder.Line.Style = borderStyle
        newBorder.Line.Weight = borderWeight
        colLeft = colLeft + colWidth
        colIndex = FindShapeIndex(slide, newBorder)
        linesArray(lineCounter) = colIndex
        lineCounter = lineCounter + 1
        
        ' For the last column add the last border on the right
        If i = numCol Then
            Set newBorder = slide.Shapes.AddLine(colLeft, rowTop, colLeft, rowTop + colHeight)
            newBorder.Line.ForeColor.RGB = borderColor
            newBorder.Line.Visible = borderVisible
            newBorder.Line.Style = borderStyle
            newBorder.Line.Weight = borderWeight
            colLeft = colLeft + colWidth
            colIndex = FindShapeIndex(slide, newBorder)
            linesArray(lineCounter) = colIndex
            lineCounter = lineCounter + 1
        End If
    Next i
        
    ' Create row borders
    colHeight = tableShape.Height
    rowWidth = tableShape.Width
    colLeft = tableShape.Left
    rowTop = tableShape.Top
    
    For j = 1 To numRows
        Set row = tbl.Rows(j)
        rowHeight = row.Height
        Set newBorder = slide.Shapes.AddLine(colLeft, rowTop, colLeft + rowWidth, rowTop)
        newBorder.Line.ForeColor.RGB = borderColor
        newBorder.Line.Visible = borderVisible
        newBorder.Line.Style = borderStyle
        newBorder.Line.Weight = borderWeight
        rowTop = rowTop + rowHeight
        RowIndex = FindShapeIndex(slide, newBorder)
        linesArray(lineCounter) = RowIndex
        lineCounter = lineCounter + 1
        
        ' For the last row add the last border on the bottom
        If j = numRows Then
            Set newBorder = slide.Shapes.AddLine(colLeft, rowTop, colLeft + rowWidth, rowTop)
            newBorder.Line.ForeColor.RGB = borderColor
            newBorder.Line.Visible = borderVisible
            newBorder.Line.Style = borderStyle
            newBorder.Line.Weight = borderWeight
            rowTop = rowTop + rowHeight
            RowIndex = FindShapeIndex(slide, newBorder)
            linesArray(lineCounter) = RowIndex
            lineCounter = lineCounter + 1
        End If
    Next j
    
    Set bordersRange = slide.Shapes.Range(linesArray)
    bordersRange.Group
    
    
    ' Delete the original table
    tableShape.Delete
    
End Sub
