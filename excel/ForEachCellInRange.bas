    Dim ColA As Range
        
    'define all cells with values in first column
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set ColA = Selection
    
    'for each cell in range
    For Each cell In ColA:
        'do this
    Next cell
