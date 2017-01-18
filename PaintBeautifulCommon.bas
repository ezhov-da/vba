Attribute VB_Name = "PaintBeautifulCommon"
'==============================================================
'Œ¡Ÿ»≈ Ã≈“Œƒ€ ƒÀﬂ Õ¿¬≈ƒ≈Õ»ﬂ  –¿—Œ“€
'==============================================================

Public Sub hideSheets(number As Integer)
    ActiveWorkbook.Sheets(number).Visible = False
End Sub

Public Sub hideIdColumn(rangeColumn As String)
    Columns(rangeColumn).EntireColumn.Hidden = True
End Sub


Public Sub setPercent(rangeColumn As String)
    Columns(rangeColumn).Style = "Percent"
End Sub


Public Sub setAllCenter(rangeColumn As String)
    Columns(rangeColumn).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Public Sub setRazryad(rangeColumn As String)
    Columns(rangeColumn).NumberFormat = "#,##"
End Sub

Public Sub setRazryadDouble(rangeColumn As String)
    Columns(rangeColumn).NumberFormat = "0.00"
End Sub


Public Sub setCodeAsText(column As Integer)
    Dim r As Long
    r = 2
    Do While (Cells(r, column) <> "")
        Cells(r, column).NumberFormat = "@"
        Cells(r, column).Value = CStr(Cells(r, column).Value)
        If Len(Cells(r, column).Value) < 10 Then
           Cells(r, column).Value = String(10 - (Len(Cells(r, column).Value)), "0") & Cells(r, column).Value
        End If
    r = r + 1
    Loop
End Sub


Public Sub changeZoom()
    ActiveWindow.Zoom = 85
End Sub

Public Sub alignmentLeft(rangeAlignment As String)
    Columns(rangeAlignment).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Public Sub alignmentCenter(rangeAlignment As String)
    Columns(rangeAlignment).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Public Sub setBold(rangeBold As String)
    Range(rangeBold).Font.Bold = True
End Sub

'Œ·˙Â‰ËÌÂÌËÂ ÒÚÓÎ·ˆÓ‚
Public Sub combine(rangeCombine As String, text As String)
    Range(rangeCombine).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = text
End Sub


Public Sub formatHeader(rowRangeSelect As String)
    Rows(rowRangeSelect).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub


Public Sub setFilter(rowFilter As String)
    Rows(rowFilter).AutoFilter
End Sub

Public Sub setFilterColumnAndValue(rowFilter As String, fieldFilter As Integer, val As String)
    Rows(rowFilter).AutoFilter field:=fieldFilter, Criteria1:=val
End Sub

Public Sub setAutoFit(rangeColumn As String)
    Columns(rangeColumn).EntireColumn.AutoFit
End Sub


Public Sub setFreezePanes(rangeRow As String)
    Rows(rangeRow).Select
    ActiveWindow.FreezePanes = True
End Sub

Public Sub setRowHeight(rangeRow As String)
    Rows(rangeRow).RowHeight = 15
End Sub

Public Sub setRowHeightRange(startRow As Long, columnId As Integer)
    Do While (Cells(startRow, columnId) <> "")
        Call setRowHeight(CStr(startRow) & ":" & CStr(startRow))
        startRow = startRow + 1
    Loop
End Sub


Public Sub groupingColumn(strRangeColumn As String)
    Columns(strRangeColumn).Columns.Group
End Sub
