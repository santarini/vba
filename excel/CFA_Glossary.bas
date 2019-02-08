Sub makeGlossarySpreadSheet()
    'define some things
    Dim i, totalRows As Integer
    Dim dataRange As Range
    'select first cell
    Range("A1").Select
    'select all cells beneath
    Set dataRange = Application.Range(Selection, Selection.End(xlDown))
    'get row count of cells beneath
    totalRows = dataRange.Rows.Count
    'for all rows in datarange
    For i = 1 To totalRows:
        'if row is odd
        If i Mod 2 = 1 Then
            'select the cell beneath it
            Range("A1").Offset(i, 0).Select
            Selection.Cut
            'move the content to the cell that is directly right and up
            ActiveCell.Offset(-1, 1).Select
            ActiveSheet.Paste
        End If
    Next i
End Sub
