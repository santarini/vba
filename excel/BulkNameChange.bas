Sub getFolderContents()

'--------------------------------------------------------------
'Author: Makoa Santarini 2019
'Takes the path of a folder listed in an Excel cell and does a drill down listing of all its containing folders and files
'To start make sure the cell containing the path of the desired folder it selected or activated
'This script will then list the containing files in rows beneath that path
'NOTE: be sure that path ends with a \
'---------------------------------------------------------------

'prevent alerts for interupting
Application.DisplayAlerts = False

'define Excel objects we'll be working with
Set MainWB = ActiveWorkbook
Set Rng = ActiveCell

'count number of elements in the path dir

Filename = Dir(Rng.Value)

i = 1
Do While Filename <> ""
    Rng.Offset(i, 0).Value = Filename
    i = i + 1
    Filename = Dir()
Loop

End Sub
Sub renameFiles()

'--------------------------------------------------------------
'Author: Makoa Santarini 2019
'Takes the output fron getFolderContents() and renames each containing object to the name directly adjacet to that listed name
'---------------------------------------------------------------

'prevent alerts for interupting
Application.DisplayAlerts = False

'define Excel objects we'll be working with
Set MainWB = ActiveWorkbook
Set Rng = ActiveCell

'define working range beneath RNG
Rng.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
RowCount = Selection.Rows.Count
Rng.Select

'go through each file and change the name of that file
For ChildFileIndex = 1 To RowCount
     OldChildFileName = Rng.Value & Rng.Offset(ChildFileIndex, 0).Value
     NewChildCileName = Rng.Value & Rng.Offset(ChildFileIndex, 1).Value
     Name OldChildFileName As NewChildCileName
Next ChildFileIndex
End Sub
