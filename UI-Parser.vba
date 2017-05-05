Sub UIParser()
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
'Loops through each cell in the range from 2 to the value of LastNonEmptyCell
For i = 2 To LastNonEmptyCell
    'Gets the value of a cell and stores in in the variable called CellValue
    CellValue = ActiveWorkbook.Worksheets("BrokenSource").Cells(i, 3).Value
    'Checks if cell is empty
    If CellValue <> "" Then
        'Checks if cell contains a line break.
        'If cell contains a line break:
        If InStr(CellValue, Chr(10)) > 0 Then
    'Prepares the string from cell to be parsed (deletes all doubled, tripled etc. New Line characters, so the string could be correctly splitted)
            CleanedString = Replace(CellValue, Chr(10) + Chr(10), Chr(10))
            'Parses the prepared string from cell
            ParsedStrings = Split(CleanedString, Chr(10))
            'Gets the length of the array that contains substrings
            ArrayLength = UBound(ParsedStrings, 1) - LBound(ParsedStrings, 1)
            'Processes each substring in the array
            For t = 0 To ArrayLength
                If ParsedStrings(t) <> "" Then
                'Assigns substring ID
                SubstringID = "!" & i & "#" & t & "!"
                'Adds substring value to the Substrings sheet
                ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 1) = ParsedStrings(t)
                'Adds substring ID to the Substrings sheet
                ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 2) = SubstringID
                'Replaces the substring with its ID in the CellValue variable
                CellValue = Replace(CellValue, ParsedStrings(t), SubstringID, 1, 1)
                RowCounter = RowCounter + 1
                End If
            Next t
            'Replaces the original cell with the cell consisting of IDs on BrokenSource sheet
            'RU
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 3) = CellValue
            'EN
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 4) = CellValue
            'UK
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 5) = CellValue
            'KK
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 6) = CellValue
            'FR
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 7) = CellValue
            'PT
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 8) = CellValue
            'ES
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 9) = CellValue
            'DE
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 10) = CellValue
            'RO
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 11) = CellValue
        'If cell doest not contain a line break:
        Else
            'Assigns substring ID
            SubstringID = "!" & i & "#0!"
            'Adds substring value to the Substrings sheet
            ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 1) = CellValue
            'Adds substring ID to the Substrings sheet
            ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 2) = SubstringID
            'Replaces the string with its ID on BrokenSource sheet
            'RU
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 3) = SubstringID
            'EN
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 4) = SubstringID
            'UK
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 5) = SubstringID
            'KK
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 6) = SubstringID
            'FR
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 7) = SubstringID
            'PT
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 8) = SubstringID
            'ES
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 9) = SubstringID
            'DE
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 10) = SubstringID
            'RO
            ActiveWorkbook.Sheets("BrokenSource").Cells(i, 11) = SubstringID
            RowCounter = RowCounter + 1
        End If
    End If
Next i
End Sub
