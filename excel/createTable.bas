Sub createTable()

Dim focus As Range
Dim tabl As Range
Dim dataRange As Range
Dim i As Long
Dim j As Long
Dim workingData As Range


'Define first table cell
Sheets("tableTest1").Activate
Sheets("tableTest1").Range("A1").Select
Set tabl = ActiveCell

'Define first Cell
Sheets("Sheet5").Activate
Sheets("Sheet5").Range("A1").Select
Set focus = ActiveCell

'Define j, last cell
focus.Select
j = Range(Selection, Selection.End(xlDown)).Rows.Count

'Define i
i = 0

Do While i < j
    'find {cheese}
    Cells.Find(What:="{cheese}", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
            
    'set focus to row just below cheese
    Set focus = ActiveCell.Offset(1, 0)
    i = ActiveCell.Offset(1, 0).Row

    'select all filled cells above cheese
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Copy
    Set workingData = Application.Selection
    
    'paste cells onto tableTest
    Sheets("tableTest1").Activate
    tabl.Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
            
    'set tabl as row below current tabl
    tabl.Select
    tabl.Offset(1, 0).Select
    Set tabl = Selection
    
    'go back to data sheet
    Sheets("Sheet5").Activate
        
    'delete the used data
    workingData.Select
    Selection.ClearContents
Loop
End Sub
