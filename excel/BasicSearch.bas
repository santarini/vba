Sub search()
    Cells.Find(What:="SEARCH-TERM", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        True, SearchFormat:=False).Activate
End Sub
