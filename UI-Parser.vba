Sub Test()
'Declare variables
Dim CellValue As String
Dim CleanedString As String
Dim ParsedStrings() As String
Dim ArrayLength As Integer
Dim SubstringID As String
Dim RowCounter As Integer
'Adds new service worksheets
ActiveWorkbook.Worksheets(1).Copy After:=Worksheets(Worksheets.Count)
ActiveWorkbook.Sheets(Worksheets.Count).Name = "BrokenSource"
ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
ActiveWorkbook.Sheets(Worksheets.Count).Name = "Substrings"
'Get last populated cell
LastPopulatedCell = ActiveWorkbook.Worksheets(2).Cells(ActiveWorkbook.Worksheets(2).Rows.Count, "B").End(xlUp).Row
For i = 2 To LastPopulatedCell
    CellValue = ActiveWorkbook.Worksheets(1).Cells(i, 3).Value
    'Checks if cell is empty
    If CellValue <> "" Then
        'Checks if cell contains a line break
        If CellValue(myString, Chr(10)) > 0 Then
            'Prepares the string from cell to be parsed
            CleanedString = Replace(CellValue, Chr(10) + Chr(10), Chr(10))
            'Parses the prepared string from cell
            ParsedStrings = Split(CleanedString, Chr(10))
            'Get the length of the array that contains substrings
            ArrayLength = UBound(ParsedStrings, 1) - LBound(ParsedStrings, 1)
            'Processes each substring in the array
            For t = 0 To ArrayLength
                If ParsedStrings(t) <> "" Then
                SubstringID = "!" & i & "#" & t & "!"
                'Add data to sheet2
                RowCounter = RowCounter + 1
                End If
            Next t
        Else
        SubstringID = "!" & i & "#0!"
        'Add data to sheet2
        RowCounter = RowCounter + 1
        End If
    End If
Next i
End Sub
