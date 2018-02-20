Sub compileByPath()

Dim FolderPath As String
Dim PathCountCondition As String
Dim FileName As String
Dim Count As Integer
Dim FileNumber As Integer
Dim MainWB As Workbook
Dim WB As Workbook

'set this workbook as the main workbook

Set MainWB = ActiveWorkbook

'define folder path
FolderPath = "C:\Users\m4k04\Desktop\vb_files\"

'count number of CSVs in folder

PathCountCondition = FolderPath & "\*.csv"

FileName = Dir(Path)

Do While FileName <> ""
   Count = Count + 1
    FileName = Dir()
Loop


For FileNumber = 1 To Count

    'open the file

    FileName = FolderPath & FileNumber & ".csv"
    
    Set WB = Workbooks.Open(FileName)
    
    'copy its contents
    
    
    
    'paste it into the main workbook
    
    
    'close file

Next i

End Sub
