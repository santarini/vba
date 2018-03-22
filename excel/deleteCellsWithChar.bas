Sub deleteCellsWithChar()

Dim Rng As Range
Dim Cell As Range
Dim CellStr As String

Set Rng = Application.InputBox(prompt:="Select tickers", Type:=8)

For Each Cell In Rng
    CellStr = Cell.Text
    If InStr(CellStr, "^") > 0 Then
    Cell.Select
    Selection.Delete Shift:=xlUp
    End If
Next Cell
End Sub
