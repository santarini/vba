'Black
'ColorIndex = 1 

'White
'ColorIndex = 2

'Red
'ColorIndex = 3 

'Green
'ColorIndex = 4 

'Blue
'ColorIndex = 5

'Yellow
'ColorIndex = 6

'Cell Color
'Selection.Interior.ColorIndex = 3

'Font Color
'Selection.Font.ColorIndex = 3


'Format("210.6", "#,##0.00")
'Result: '210.60'

'Format("210.6", "Standard")
'Result: '210.60'

'Format("0.981", "Percent")
'Result: '98.10%'

'Format("1267.5", "Currency")
'Result: '$1,267.50'

'Format("Sep 3, 2003", "Short Date")
'Result: '9/3/2003'

returnOnEquity = -0.0123

Selection.Value = Format(returnOnEquity, "Percent")
If Selection > 0 Then
    Selection.Font.ColorIndex = 4
End If
If Selection < 0 Then
    Selection.Font.ColorIndex = 3
End If
