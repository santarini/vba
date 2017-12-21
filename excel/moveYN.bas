Sub moveYN()

Dim focus As Range
Dim i As Long
Dim j As Long

Range("A1").Select
j = Range(Selection, Selection.End(xlDown)).Rows.Count

Range("I1").Select
Set focus = ActiveCell

i = 0

Do While i < j
    If InStr(1, (focus.Value), "Yes") > 0 Then
        'move data two cells over
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        'set focus to next cell
        i = i + 1
        focus.Offset(1, -2).Select
        Set focus = ActiveCell
    ElseIf InStr(1, (focus.Value), "No") > 0 Then
        'move data two cells over
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        'set focus to next cell
        i = i + 1
        focus.Offset(1, -2).Select
        Set focus = ActiveCell
    Else
        'set focus to next cell
        focus.Offset(1, 0).Select
        Set focus = ActiveCell
        'increment i
        i = i + 1
    End If
Loop
End Sub
