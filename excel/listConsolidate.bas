Sub createMasterList()

Dim ws As Worksheet
Dim wsCount As Integer


'create master page
Sheets.Add.Name = "Master"

Worksheets("Master").Activate
Range("A1").Value = "Status"
Range("B1").Value = "Email"
Range("C1").Value = "Group"
Range("D1").Value = "First Name"
Range("E1").Value = "Last Name"
Range("A2").Select

'go to first sheet
Sheets(2).Select

wsCount = ThisWorkbook.Sheets.Count

'for each sheet in workbook

For i = 2 To wsCount:
    Sheets(i).Activate
    Range("A2:E2").Select
    
        'select down
    Range(Selection, Selection.End(xlDown)).Select
    
        'copy selection
    Selection.Copy
    
        
        'navigate to master page
        
    Worksheets("Master").Activate
    
        'move to bottom of content
        
    
        'paste content
    Worksheets("Master").Paste
    
        'reset source cell
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    
    Next
End Sub
