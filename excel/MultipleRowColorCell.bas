Sub colorCells()

'collin is super lame

Dim rng, cell, indexCell, WorkingRow As Range
Dim str As String
Dim i As Integer

Set rng = Selection

For Each indexCell In rng
indexCell.Select

c = ActiveCell.Interior.Color
i = indexCell.Row

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
Next indexCell

rng.Select

End Sub
