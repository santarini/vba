Sub compileByPath()

Dim FolderPath As String
Dim PathCountCondition As String
Dim FileName As String
Dim Count As Integer
Dim FileNumber As Integer
Dim MainWB As Workbook
Dim WB As Workbook
Dim Rng As Range

'set this workbook as the main workbook

Set MainWB = ActiveWorkbook
MainWB.Sheets.Add.Name = "DataSet"
Set Rng = Range("A1")


'define folder path
FolderPath = "C:\Users\m4k04\Desktop\vb_files\"

'count number of CSVs in folder

PathCountCondition = FolderPath & "\*.csv"

FileName = Dir(PathCountCondition)

Do While FileName <> ""
   Count = Count + 1
    FileName = Dir()
Loop

For FileNumber = 1 To Count

    'open the file

    FileName = FolderPath & FileNumber & ".csv"
    
    Set WB = Workbooks.Open(FileName)
    
    'copy its contents
    
    WB.Activate

    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'paste it into the main workbook
    
    MainWB.Activate
    Worksheets("DataSet").Activate
    Rng.Select
    ActiveSheet.Paste
    Rng.Offset(0, 1).Select
    Set Rng = ActiveCell
    
    'close file
    WB.Close
    

Next FileNumber
End Sub
