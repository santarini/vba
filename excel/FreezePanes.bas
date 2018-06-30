'freeze panes
Range("A1").Select
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 1
    End With
ActiveWindow.FreezePanes = True
