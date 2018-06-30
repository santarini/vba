Public Function

On Error GoTo Handler:
  '
  '
  '

ExitErrorPoint:

Exit Function
'Exit function occurs before the handler because if there isn't an exit, the code will read through the handler normally as if it were the next logical piece of code

Handler:
  '
  '
  '
  '
Resume ExitErrorPoint
                      
End Function

