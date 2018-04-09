Sub HTTPrequest()
    Dim objHTTP As Object
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim URL As String
    URL = "http://www.google.com"
    
    Dim sendResult As String

    objHTTP.Open "GET", URL, False
    objHTTP.Send
    sendResult = objHTTP.ResponseText
    MsgBox sendResult
End Sub
