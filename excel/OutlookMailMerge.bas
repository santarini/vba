Sub Open_Outlook()

Dim objOutlook As Object
Dim Msg As Object

'there can be no open instances of Outlook while this is running
Set objOutlook = CreateObject("Outlook.Application")
Set Msg = objOutlook.CreateItem(0)

With Msg
 .To = "someone@something.com;"
 .Subject = "What are you sending this for"
 .HTMLBody "Whatever you whant to say"
 .send
End With

End Sub
