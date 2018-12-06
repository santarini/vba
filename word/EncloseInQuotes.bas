'Enclose selection in quotes
    With Selection.Range
        .Text = Chr(34) & .Text & Chr(34)
    End With
