Sub SearchBoldText()
    Dim rng As Range
    Set rng = ThisDocument.Range(0, 0)
    With rng.Find
        .ClearFormatting
        .Format = True
        .Font.Bold = True
        While .Execute
            rng.Select
            With Selection
                .Font.Bold = False
                .Font.Underline = True
                .InsertParagraphBefore
                .InsertAfter ","
            End With
        Wend
    End With
    Set rng = Nothing
End Sub
