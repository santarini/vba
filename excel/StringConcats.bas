'Takes the first inital and the last word in a cell
TempStartValue = Range("H2").Value

MsgBox TempStartValue

FinalValue1 = Left(TempStartValue, 1)
FinalValue2 = Split(TempStartValue, " ")
MsgBox FinalValue1
MsgBox FinalValue2(UBound(FinalValue2))
