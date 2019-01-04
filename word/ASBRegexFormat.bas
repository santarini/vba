Sub ASBRegexFormat()
    Dim strTemp As String
    Dim lastRow As Long
    lastRow = ActiveDocument.BuiltInDocumentProperties("Number Of Lines")

'isolate dates  and create seperate lines

    Selection.EndKey Unit:=wdLine
    'Find instance with regex and replace
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
    
    Selection.Delete Unit:=wdCharacter, Count:=1
    
    'reset line count
    lastRow = ActiveDocument.BuiltInDocumentProperties("Number Of Lines")
    MsgBox lastRow
    
'isolate massive numbers
End Sub
