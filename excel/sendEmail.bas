Sub sendEmail()
Dim Edge As Object

Set Edge = CreateObject("InternetExplorer.Application")

IE.Top = 0
IE.Left = 0
IE.Width = 800
IE.Height = 600
IE.AddressBar = 0
IE.StatusBar = 0
IE.Toolbar = 0
IE.Visible = True
    
Edge.Navigate ("https://mail.google.com/mail/u/1/#inbox")
Do
DoEvents
Loop Until IE.ReadyState = 4

'If ReadyState is workign use an actual wait value
    'Application.Wait Now + TimeValue("00:00:05")
'If you need to quit IE
    'IE.Quit

'set object at HTML element
'interact with HTML elements








End Sub
