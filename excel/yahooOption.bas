Public Sub parsehtml()
Dim http As Object
Dim html As New HTMLDocument
Dim yahooHeader, priceParentDiv, priceSubParentDiv, priceSubSubParentDiv, optionSpan As Object
', titleElem As Object, detailsElem As Object, topic As HTMLHtmlElement
Dim i As Integer
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://finance.yahoo.com/quote/MSFT210618C00130000?p=MSFT210618C00130000", False
http.send
html.body.innerHTML = http.responseText

Set yahooHeader = html.getElementById("quote-header-info")

MsgBox yahooHeader.innerText

Set priceParentDiv = yahooHeader.getElementsByTagName("td")(3)
Set priceSubParentDiv = priceParentDiv.getElementByTagName("div")(1)
Set priceSubSubParentDiv = priceSubParentDiv.getElementsByTagName("div")(1)
Set optionSpan = priceSubSubParentDiv.getElementsByTagName("span")(1)

'MsgBox yahooHeader.getElementsByTagName("div")(1).innerText
'MsgBox yahooHeader.getElementsByTagName("div")(2).innerText
MsgBox optionSpan.innerText


'Sheets(1).Cells(i, 1).Value = titleElem.getElementsByTagName("a")(0).innerText
'Sheets(1).Cells(i, 2).Value = titleElem.getElementsByTagName("a")(0).href
'Set detailsElem = topic.NextSibling.getElementsByTagName("td")(1)
'Sheets(1).Cells(i, 3).Value = detailsElem.getElementsByTagName("span")(0).innerText
'Sheets(1).Cells(i, 4).Value = detailsElem.getElementsByTagName("a")(0).innerText
'i = i + 1
'Next
End Sub
