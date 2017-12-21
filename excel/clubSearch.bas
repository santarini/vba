Sub clubSearch()

Dim focus As Range
Dim i As Long
Dim j As Long

Range("A1").Select
j = Range(Selection, Selection.End(xlDown)).Rows.Count

Range("C1").Select
Set focus = ActiveCell


i = 1

Do While i < j
    If InStr(1, (focus.Value), "Farms") > 0 Then
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        i = i + 1
        focus.Offset(1, -1).Select
        Set focus = ActiveCell
    Else
        i = i + 1
        focus.Offset(1, 0).Select
        Set focus = ActiveCell
    End If
Loop

End Sub
