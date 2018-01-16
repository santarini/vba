Sub example()

Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
MyRequest.Open "GET", "https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=MSFT&interval=1min&apikey=demo"
MyRequest.Send

Dim Json As Object
Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)
Dim Dict As New Dictionary
Dict.CompareMode = CompareMethod.TextCompare
Dict("A") = "Stock Quotes"

LastRefreshed = Json(Dict.Item("A"))(2)("1. symbol")
MsgBox LastRefreshed

End Sub
