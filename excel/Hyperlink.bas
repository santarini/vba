'Add External Hyperlink
ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="http://makoasantarini.com/", TextToDisplay:="Click Here"

'Add External Email Hyperlink
ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="mailto:someone@example.com?subject=hello", ScreenTip:="Send an email", TextToDisplay:="Create Email"

'Add Internal Hyperlink
ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="CtoCCOV!A1", TextToDisplay:="Click Here"
