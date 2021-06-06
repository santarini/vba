Option Explicit
Sub TOS_Raw_OptionHacker_Format()
    
    ''''Start Clean Up
    
    Range("A1").Select
    
    '''Formatting
    'Delete the top two rows
    
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    
    'Auto adjust col width
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    'Hard code values
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Bold B & W the top row
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .Bold = True
    End With
    
    '''filter out Symbols with "./"
    
    Dim rng As Range, i As Integer
        
    'define all cells with values in first column
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set rng = Selection
    
    'for each cell in range
    For i = rng.Rows.Count To 1 Step -1
        If InStr(1, rng.Cells(i), "./") > 0 Then
            rng.Cells(i).EntireRow.Delete
        End If
    Next i
    
    '''filter out descriptions with fractional option contracts
        
    'define all cells with values in first column
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set rng = Selection
    
    'for each cell in range
    For i = rng.Rows.Count To 1 Step -1
        If InStr(1, rng.Cells(i), "/") > 0 Then
            rng.Cells(i).EntireRow.Delete
        End If
    Next i
    
    'Finish Clean Up
    Range("A1").Select
End Sub

