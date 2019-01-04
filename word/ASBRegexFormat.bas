Sub ASBRegexFormat()
    Dim strTemp As String
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
        strTemp = ""
End Sub
