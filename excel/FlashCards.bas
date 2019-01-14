Sub FlashCards()
    'Select all the cells
    Cells.Select
    'Resize the cells
    Selection.ColumnWidth = 30
    Selection.RowHeight = 172
    'Slap some borders on them cells
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    'center the text horizontally and vertically
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    'set print area
    Range("A1:C4").Select
    Application.CutCopyMode = False
    ActiveSheet.PageSetup.PrintArea = "$A$1:$C$4"
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    Range("A1:C8").Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$C$8"
    Range("A1:F8").Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$F$8"
    Range("A1:I8").Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$I$8"
    'reset selection
    Range("A1").Select
End Sub
