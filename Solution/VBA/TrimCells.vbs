Sub TrimCells()
    
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    
    For i = 1 To lastRow
        For j = 1 To lastColumn
            Cells(i, j).Value = Trim(Cells(i, j).Value)
        Next j
    Next i
    
End Sub
