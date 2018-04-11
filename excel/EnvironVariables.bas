Sub ListEnvironVariables()
    Dim strEnviron As String
    Dim i As Long
    For i = 1 To 255
        strEnviron = Environ(i)
        If LenB(strEnviron) = 0& Then Exit For
        Debug.Print strEnviron
    Next
End Sub
