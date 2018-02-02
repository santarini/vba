Sub countInCell()

Dim rngData As Range
Dim rngBgn As Range
Dim rngEnd As Range
Dim rowCount As Integer
Dim currentCell As Range
Dim i As Integer

Set rngData = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
rowCount = rngData.Rows.Count

Set rngBgn = rngData.Cells(1, 1)

Sheets.Add.Name = "temp"

Set rngEnd = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)

rngEnd.Value = rngBgn

For i = 1 To rowCount Step 1

If rngBgn <> rngEnd Then
Worksheets("temp").Activate
rngEnd.Offset(1, 0).Select
Set rngEnd = ActiveCell
rngEnd.Value = rngBgn
End If

If rngBgn = rngEnd Then
Worksheets("test").Activate
rngBgn.Offset(0, 1).Select
    
    If rngBgn.Offset(0, 2).Value <> "" Then
    Range(Selection, Selection.End(xlToRight)).Select
    
    End If
Selection.Copy
Worksheets("temp").Activate
    
    If rngEnd.Offset(0, 1).Value <> "" Then
    rngEnd.Select
    Selection.End(xlToRight).Select
    End If

Selection.Offset(0, 1).Select
ActiveSheet.Paste
Worksheets("test").Activate
rngBgn.Offset(1, 0).Select
Set rngBgn = ActiveCell
End If
Next

End Sub
