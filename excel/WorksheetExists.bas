Dim Sht As Worksheet
Dim WorksheetExists As Boolean
Dim TestName As True

TestName = "Sheet1"

WorksheetExists = False
For Each Sht In ThisWorkbook.Worksheets
    If Sht.Name = TestName Then
        WorksheetExists = True
    End If
Next Sht

If (WorksheetExists = False) Then
  'action
End If

If (WorksheetExists = True) Then
  'action
End If
