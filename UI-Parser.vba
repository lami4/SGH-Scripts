Sub Test()
'Declare variables
Dim CellValue As String
Dim CleanedString As String
Dim ParsedStrings() As String
Dim ArrayLength As Integer
Dim SubstringID As String
Dim RowCounter As Integer
RowCounter = 1
'Adds new service worksheets
ActiveWorkbook.Worksheets(1).Copy After:=Worksheets(Worksheets.Count)
ActiveWorkbook.Sheets(Worksheets.Count).Name = "BrokenSource"
ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
ActiveWorkbook.Sheets(Worksheets.Count).Name = "Substrings"
    'Gets non-empty cell in column C
    LastNonEmptyCell = ActiveWorkbook.Worksheets("BrokenSource").Cells(ActiveWorkbook.Worksheets("BrokenSource").Rows.Count, "C").End(xlUp).Row
'Loops through each non-empty cell in column C
For i = 2 To LastNonEmptyCell
    CellValue = ActiveWorkbook.Worksheets("BrokenSource").Cells(i, 3).Value
    'Checks if cell is empty
    If CellValue <> "" Then
        'Checks if cell contains a line break
        If InStr(CellValue, Chr(10)) > 0 Then
            'Prepares the string from cell to be parsed
            CleanedString = Replace(CellValue, Chr(10) + Chr(10), Chr(10))
            'Parses the prepared string from cell
            ParsedStrings = Split(CleanedString, Chr(10))
            'Get the length of the array that contains substrings
            ArrayLength = UBound(ParsedStrings, 1) - LBound(ParsedStrings, 1)
            'Processes each substring in the array
            For t = 0 To ArrayLength
                If ParsedStrings(t) <> "" Then
                'Assigns substring ID
                SubstringID = "!" & i & "#" & t & "!"
                'Adds substring value to the Substrings sheet
                ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 1) = ParsedStrings(t)
                'Adds ID to the Substrings sheet
                ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 2) = SubstringID
                'Replaces the substring with its ID in the CellValue variable
                CellValue = Replace(CellValue, ParsedStrings(t), SubstringID, 1, 1)
                RowCounter = RowCounter + 1
                End If
            Next t
            'Replaces the original cell with the indexed cell
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 3) = CellValue
        Else
            'Assigns substring ID
            SubstringID = "!" & i & "#0!"
            'Adds substring value to the Substrings sheet
            ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 1) = CellValue
            'Adds ID to the Substrings sheet
            ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 2) = SubstringID
            'Replaces the string with its ID
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 3) = SubstringID
            RowCounter = RowCounter + 1
        End If
    End If
Next i
End Sub
