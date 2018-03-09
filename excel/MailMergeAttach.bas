Sub MailMerge()

Dim NameHeader As Range
Dim AdrsHeader As Range
Dim NameRange As Range
Dim NameCount As Integer
Dim AdrsRange As Range
Dim AdrsCount As Integer
Dim FileCellCount As Integer

Dim i As Integer

Dim objOutlook As Object
Dim Msg As Object
Dim Body1 As String
Dim Body2 As String
Dim Body3 As String
Dim Body4 As String
Dim Signature As String
Dim FileCellHeader As Range
Dim FileCellRange As Range
Dim FileCell As Range

    Cells.Find(What:="First Name", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Activate
    Set NameHeader = Selection
    NameHeader.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set NameRange = Selection
    NameCount = NameRange.Rows.Count
    
    Cells.Find(What:="Email", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Activate
    Set AdrsHeader = Selection
    AdrsHeader.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set AdrsRange = Selection
    AdrsCount = AdrsRange.Rows.Count
    
    Cells.Find(What:="Attachment", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Activate
    Set FileCellHeader = Selection
    FileCellHeader.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set FileCellRange = Selection
    FileCellCount = FileCellRange.Rows.Count
    
    Body1 = "Paragraph1"
    Body2 = "Paragraph 2"
    Body3 = "Paragraph 3"
    Body4 = "Paragraph 4"
    Signature = "<img src='C:\Users\m4k04\Desktop\logo.png'><br><b>Makoa Santarini</b><br><font color='blue'>Principal <br>Makoa Systems, Inc. <br>Address 1<br>Address 2 <br>City, State, Zip <br>Phone 1<br>Phone 2 <br>Email<br></font>http://www.makoasystems.com"
    
For i = 1 To NameCount

    NameRange.Rows(i).Select
    FirstName = Selection
    AdrsRange.Rows(i).Select
    EmailAdrs = Selection
    FileCellRange.Rows(i).Select
    Set FileCell = Selection
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set Msg = objOutlook.CreateItem(0)
    
    With Msg
     .to = EmailAdrs
     .Subject = "Join us on Mar. 29th at 5 pm for a reception with LG Doug Chin - The Pacific Club"
     .HTMLBody = "Aloha " & FirstName & ",<br><br>" & Body1 & "<br><br>" & Body2 & "<br><br>" & Body3 & "<br><br>" & Body4 & "<br><br><br>" & "Mahalo,<br>" & Signature
        If Trim(FileCell) <> "" Then
            If Dir(FileCell.Value) <> "" Then
                .Attachments.Add FileCell.Value
            End If
        End If
     .send
    End With

Next i

End Sub
