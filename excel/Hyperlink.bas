'Add External Hyperlink
ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="http://makoasantarini.com/", TextToDisplay:="Click Here"


'Add Internal Hyperlink
ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="CtoCCOV!A1", TextToDisplay:="Click Here"
