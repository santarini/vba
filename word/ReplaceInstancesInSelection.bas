'Replace spaces in selection with %
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = "%"
        .Forward = True
        .Wrap = wdFindStop 'wdFindStop is key to making it only search your highligted selection, avoid wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
