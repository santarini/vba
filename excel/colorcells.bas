Sub colorCells()

Dim rng, cell, WorkingRow As Range
Dim str As String
Dim i As Integer

c = ActiveCell.Interior.Color

Set rng = Selection
i = rng.Row

R = c Mod 256
G = c \ 256 Mod 256
B = c \ 65536 Mod 256

Rows(i & ":" & i).Select

Set WorkingRow = Selection

For Each cell In WorkingRow

If cell = "X" Or cell = "x" Then
    cell.Interior.Color = RGB(R, G, B)
End If
Next cell

rng.Select

End Sub
