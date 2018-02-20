Sub compileByPath()

Dim PathName As String
Dim FileName As String
Dim FileNumber As Integer
Dim WB As Workbook

FileNumber = 1

FilePath = "C:\Users\m4k04\Desktop\vb_files"
FileName = FilePath & FileNumber & ".csv"

Set WB = Workbooks.Open(FileName)

End Sub
