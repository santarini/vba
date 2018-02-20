Sub compileByPath()

Dim FolderPath As String
Dim FileName As String
Dim FileNumber As Integer
Dim WB As Workbook

FileNumber = 2

FolderPath = "C:\Users\m4k04\Desktop\vb_files\"
FileName = FolderPath & FileNumber & ".csv"


Set WB = Workbooks.Open(FileName)

End Sub
