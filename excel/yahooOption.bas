Public Sub yahooOptionQuote()
Dim http As Object
Dim html As New HTMLDocument
Dim yahooHeader As Object, priceParentDiv, optionSpan As Object
Dim i As Integer
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://finance.yahoo.com/quote/MSFT210618C00130000?p=MSFT210618C00130000", False
http.send
html.body.innerHTML = http.responseText

Set yahooHeader = html.getElementById("quote-header-info")
Set priceParentDiv = yahooHeader.getElementsByTagName("div")(7)
Set optionSpan = priceParentDiv.getElementsByTagName("span")(0)
MsgBox optionSpan.innerText

End Sub
