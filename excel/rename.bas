Sub RenameFiles()

Dim fileBeforeRng As Range
Dim fileAfterRng As Range
Dim i As Integer
Dim j As Integer

Set fileBeforeRng = Application.InputBox(prompt:="Where is the first starting value?", Type:=8)
Set fileAfterRng = Application.InputBox(prompt:="Where is the first ending value?", Type:=8)

fileBeforeRng.Select
Range(Selection, Selection.End(xlDown)).Select

j = Selection.Rows.Count
MsgBox j
i = 1

Do While i <= j
   Name fileBeforeRng As fileAfterRng
    fileBeforeRng.Offset(1, 0).Select
    Set fileBeforeRng = Selection
    fileAfterRng.Offset(1, 0).Select
    Set fileAfterRng = Selection
    i = i + 1
Loop
End Sub
