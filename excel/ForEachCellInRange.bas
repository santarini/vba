    Dim rng As Range, i As Integer
        
    'define all cells with values in first column
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set rng = Selection
    
    'for each cell in range (start from bottom and step backward)
    For i = rng.Rows.Count To 1 Step -1
        If InStr(1, rng.Cells(i), "./") > 0 Then
            rng.Cells(i).EntireRow.Delete
        End If
    Next i
