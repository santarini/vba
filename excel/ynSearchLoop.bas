Sub ynSearchLoop()

Dim focus As Range
Dim i As Long
Dim j As Long


Sheets("Sheet5").Range("A1").Select
Set focus = ActiveCell

i = 0
j = 46000

Do While i < j
    If focus = "Yes" Then
        'insert a row
        focus.Offset(1, 0).EntireRow.Insert
        'insert unqiue phrase
        focus.Offset(1, 0).Value = "{cheese}"
        'set focus to next cell
        focus.Offset(1, 0).Select
        Set focus = ActiveCell
        'increment i
        i = i + 1
    
    ElseIf focus = "No" Then
        'insert a row
        focus.Offset(1, 0).EntireRow.Insert
        'insert unqiue phrase
        focus.Offset(1, 0).Value = "{cheese}"
        'set focus to next cell
        focus.Offset(1, 0).Select
        Set focus = ActiveCell
        'increment i
        i = i + 1
    
    Else
        'set focus to next cell
        focus.Offset(1, 0).Select
        Set focus = ActiveCell
        'increment i
        i = i + 1
    End If
Loop


End Sub
