Sub elapsedTime()
Dim StartTime As Double
Dim SecondsElapsed As Double
    StartTime = Timer
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
End Sub
