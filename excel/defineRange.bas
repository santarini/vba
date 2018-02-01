'define some ranges

Dim rng As Range
Dim rngTrgt As Range
Dim currentCell As Range

'set those ranges

Set rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
Set rngTrgt = rng.Cells(1, 1)
rngTrgt.Offset(0, 1).Select
Set currentCell = ActiveCell

'define some ints
    
Dim rowCount As Integer
Dim cellCount As Integer
    
'store the row and cell count of the range
    
rowCount = rng.Rows.Count
cellCount = rng.Cells.Count

'verify
        
MsgBox rowCount
MsgBox cellCount
MsgBox rngTrgt
MsgBox currenCell
