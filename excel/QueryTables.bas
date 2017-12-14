Dim adrs As String

adrs = "https://www.google.com"

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Sheets.Add.Name = "Temp"

QueryQuote:
        With Sheets("Temp").QueryTables.Add(Connection:="URL;" & adrs, Destination:=ActiveSheet.Range("a1"))
            .BackgroundQuery = True
            .TablesOnlyFromHTML = True
            .Refresh BackgroundQuery:=False
            .SaveData = True
        End With
