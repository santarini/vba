Sub UniqueNames()

Dim cellCount As Integer
Dim uniqName As New Collection, a
Dim allNames() As Variant
Dim i As Long

Range("A1").Select
Set nameCol = Range(Selection, Selection.End(xlDown))
cellCount = nameCol.Rows.Count

ReDim allNames(1 To cellCount) As Variant

i = 1
For Each cell In nameCol
    allNames(i) = cell.Value
i = i + 1
Next cell

On Error Resume Next
For Each a In allNames
   uniqName.Add a, a
Next

For i = 1 To uniqName.Count
    Cells(i, 2) = uniqName(i)
Next

End Sub
