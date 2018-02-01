Dim rng As Range
Dim rowCount As Integer
Dim cellCount As Integer


Set rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
rowCount = rng.Rows.Count
cellCount = rng.Cells.Count

MsgBox rowCount
MsgBox cellCount
