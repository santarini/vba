Sub ASBFormatRegex()
'find dates ##-## ##-##
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "([0-9]{2})-([0-9]{2}) ([0-9]{2})-([0-9]{2})"
        .InsertBefore Chr(13)
        .InsertAfter ","
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.InsertBefore Chr(13)
    Selection.InsertAfter ","
    
'find dollar values

End Sub
Sub WildcardTest()
    Dim strTemp As String
    Dim lastRow As Long
    lastRow = ActiveDocument.BuiltInDocumentProperties("Number Of Lines")
    MsgBox lastRow

'script currently needs to be run when cursor is at the top of the page

'############# isolate dates  and create seperate lines
    'ActiveDocument.Range(0, 0).Select
    'Selection.HomeKey Unit:=wdCharacter
    'Selection.EndKey Unit:=wdLine
    
'Find dates and put line breaks infront of them
    For i = 1 To lastRow
        Selection.Find.ClearFormatting
        With Selection.Find
            .MatchWildcards = True
            .Text = "([0-9]{2})-([0-9]{2}) ([0-9]{2})-([0-9]{2})"
            .Forward = False
            .Execute
            strTemp = Selection
            .Wrap = wdFindContinue
            .Replacement.Text = Chr(13) + strTemp + ","
            .Execute Replace:=wdReplaceOne
        End With
        'clear strTemp
        strTemp = ""
    'Next instance
    Next i
    
    'Delete extra top line
    'This creates an error delteing a random character if script is launched with cursor in middle of text
    Selection.Delete Unit:=wdCharacter, Count:=1
    
    'reset line count
    lastRow = ActiveDocument.BuiltInDocumentProperties("Number Of Lines")
       
'isolate dollar amounts
    For i = 1 To lastRow
        Selection.Find.ClearFormatting
        With Selection.Find
            .MatchWildcards = True
            .Text = "$"
            .Forward = False
            .Execute
            strTemp = Selection
            .Wrap = wdFindContinue
            .Replacement.Text = ", $"
            .Execute Replace:=wdReplaceOne
        End With
        'clear strTemp
        strTemp = ""
    'Next instance
    Next i
    
    
'Delete massive numbers
    'Selection.HomeKey Unit:=wdLine
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "([0-9 ]{25})"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

 
'Delte straggler numbers
    'Selection.HomeKey Unit:=wdLine
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "([ ])([0-9])([ ])"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'reset cursor
    
'Find first dates and put commas after them
    For i = 1 To lastRow
        Selection.Find.ClearFormatting
        With Selection.Find
            .MatchWildcards = True
            .Text = "([0-9]{2})-([0-9]{2})([ ])"
            .Forward = False
            .Execute
            strTemp = Selection
            .Wrap = wdFindContinue
            .Replacement.Text = strTemp + ","
            .Execute Replace:=wdReplaceOne
        End With
        'clear strTemp
        strTemp = ""
    'Next instance
    Next i
    
End Sub
