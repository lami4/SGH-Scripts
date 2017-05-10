Sub UIParserAndCellReference()
'Declare variables
Dim CellValue As String
Dim IDCellValue As String
Dim PreparedString As String
Dim Substrings() As String
Dim ArrayLength As Integer
Dim SubstringID As String
Dim RowCounter As Integer
Dim SubstringNumber As Integer
Dim CellReference As String
Dim FirstTwoCharacters As String
Dim FormulaRU As String
Dim FormulaEN As String
Dim FormulaUK As String
Dim FormulaKK As String
Dim FormulaFR As String
Dim FormulaPT As String
Dim FormulaES As String
Dim FormulaDE As String
Dim FormulaRO As String
RowCounter = 2
SubstringNumber = 1
'Adds new service worksheets
ActiveWorkbook.Worksheets(1).Copy After:=Worksheets(Worksheets.Count)
ActiveWorkbook.Sheets(Worksheets.Count).Name = "CellReferenceParsing"
ActiveWorkbook.Worksheets(1).Copy After:=Worksheets(Worksheets.Count)
ActiveWorkbook.Sheets(Worksheets.Count).Name = "UIDParsing"
ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
ActiveWorkbook.Sheets(Worksheets.Count).Name = "Substrings"
'Adds headers to Substrings sheet
ActiveWorkbook.Sheets("Substrings").Cells(1, 1) = "Service ID"
ActiveWorkbook.Sheets("Substrings").Cells(1, 2) = "Native ID"
ActiveWorkbook.Sheets("Substrings").Cells(1, 3) = "RU"
ActiveWorkbook.Sheets("Substrings").Cells(1, 4) = "EN"
ActiveWorkbook.Sheets("Substrings").Cells(1, 5) = "UK"
ActiveWorkbook.Sheets("Substrings").Cells(1, 6) = "KK"
ActiveWorkbook.Sheets("Substrings").Cells(1, 7) = "FR"
ActiveWorkbook.Sheets("Substrings").Cells(1, 8) = "PT"
ActiveWorkbook.Sheets("Substrings").Cells(1, 9) = "ES"
ActiveWorkbook.Sheets("Substrings").Cells(1, 10) = "DE"
ActiveWorkbook.Sheets("Substrings").Cells(1, 11) = "RO"
    'Gets last non-empty cell in column C
    LastNonEmptyCell = ActiveWorkbook.Worksheets("CellReferenceParsing").Cells(ActiveWorkbook.Worksheets("CellReferenceParsing").Rows.Count, "C").End(xlUp).Row
'Loops through each cell in the range from 2 to the value of LastNonEmptyCell
For i = 2 To LastNonEmptyCell
    'Gets the value of a cell and stores in in the variable called CellValue
    CellValue = ActiveWorkbook.Worksheets("CellReferenceParsing").Cells(i, 3).Value
    IDCellValue = CellValue
    'Checks if cell is empty
    If CellValue <> "" Then
        'Checks if cell contains a line break.
        'If cell contains a line break:
        If InStr(CellValue, Chr(10)) > 0 Then
            'Prepares the string from cell to be parsed (deletes all doubled, tripled etc. New Line characters, so the string could be correctly splitted)
            PreparedString = Replace(CellValue, Chr(10) + Chr(10), Chr(10))
            'Splits prepared string by New Line charachter into substrings
            Substrings = Split(PreparedString, Chr(10))
            'Gets the length of the array that contains substrings
            ArrayLength = UBound(Substrings, 1) - LBound(Substrings, 1)
            'Processes each substring in the array
            For t = 0 To ArrayLength
                If Substrings(t) <> "" Then
                'Assigns substring ID
                SubstringID = "!" & i & "#" & SubstringNumber & "!"
                'Adds substring value to the Substrings sheet
                ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 3) = Substrings(t)
                'Adds substring ID to the Substrings sheet
                ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 1) = SubstringID
                'Adds substring native ID (File + Key + Index in array) to the Substrings sheet
                ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 2) = ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 1).Value & "/" & ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 2).Value & "/" & SubstringNumber
                'Creates a cell reference to a substring on Substrings sheet
                CellReference = " & Substrings!C" & RowCounter
                'Adds a cell reference to a substring on Substrings sheet to the cell being processed on CellReferenceParsing sheet
                CellValue = Replace(CellValue, Substrings(t), CellReference, 1, 1)
                'Replaces the substring with its ID in the CellValue variable
                IDCellValue = Replace(IDCellValue, Substrings(t), SubstringID, 1, 1)
                'Incriminates RowCounter variable by 1
                RowCounter = RowCounter + 1
                SubstringNumber = SubstringNumber + 1
                End If
            Next t
            'After all reference were added to the cell being processed on CellReferenceParsing sheet, macro replaces all New Line characters with the " & СИМВОЛ(10)" string
            CellValue = Replace(CellValue, Chr(10), " & СИМВОЛ(10)")
            'If the first two characters in the cell being processed on CellReferenceParsing sheet are " &", macro repalces it with "=" to create a formula
            FirstTwoCharacters = Left(CellValue, 2)
            If FirstTwoCharacters = " &" Then
            CellValue = Replace(CellValue, FirstTwoCharacters, "=", 1, 1)
            End If
            'Replaces the original cell on CellReferenceParsing sheet with the formula that contains cell references to substrings and replaces the original cell on UIDParsing sheet with the cell cinsisting of UIDs
            'RU
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 3).FormulaLocal = CellValue
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 3) = IDCellValue
            'EN
            FormulaEN = Replace(CellValue, "Substrings!C", "Substrings!D")
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 4).FormulaLocal = FormulaEN
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 4) = IDCellValue
            'UK
            FormulaUK = Replace(CellValue, "Substrings!C", "Substrings!E")
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 5).FormulaLocal = FormulaUK
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 5) = IDCellValue
            'KK
            FormulaKK = Replace(CellValue, "Substrings!C", "Substrings!F")
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 6).FormulaLocal = FormulaKK
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 6) = IDCellValue
            'FR
            FormulaFR = Replace(CellValue, "Substrings!C", "Substrings!G")
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 7).FormulaLocal = FormulaFR
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 7) = IDCellValue
            'PT
            FormulaPT = Replace(CellValue, "Substrings!C", "Substrings!H")
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 8).FormulaLocal = FormulaPT
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 8) = IDCellValue
            'ES
            FormulaES = Replace(CellValue, "Substrings!C", "Substrings!I")
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 9).FormulaLocal = FormulaES
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 9) = IDCellValue
            'DE
            FormulaDE = Replace(CellValue, "Substrings!C", "Substrings!J")
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 10).FormulaLocal = FormulaDE
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 10) = IDCellValue
            'RO
            FormulaRO = Replace(CellValue, "Substrings!C", "Substrings!K")
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 11).FormulaLocal = FormulaRO
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 11) = IDCellValue
            'Sets SubstringNumber bac to 1
            SubstringNumber = 1
        'If cell doest not contain a line break:
        Else
            'Assigns substring ID
            SubstringID = "!" & i & "#0!"
            'Adds substring value to the Substrings sheet
            ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 3) = CellValue
            'Adds substring ID to the Substrings sheet
            ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 1) = SubstringID
            'Adds substring native ID (File + Key + ) to the Substrings sheet
            ActiveWorkbook.Sheets("Substrings").Cells(RowCounter, 2) = ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 1).Value & "/" & ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 2).Value & "/0"
            'Replaces the original cell on CellReferenceParsing sheet with the formula that contains cell reference to substring
            'RU
            FormulaRU = "= Substrings!C" & RowCounter
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 3).FormulaLocal = FormulaRU
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 3) = SubstringID
            'EN
            FormulaEN = "= Substrings!D" & RowCounter
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 4).FormulaLocal = FormulaEN
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 4) = SubstringID
            'UK
            FormulaUK = "= Substrings!E" & RowCounter
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 5).FormulaLocal = FormulaUK
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 5) = SubstringID
            'KK
            FormulaKK = "= Substrings!F" & RowCounter
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 6).FormulaLocal = FormulaKK
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 6) = SubstringID
            'FR
            FormulaFR = "= Substrings!G" & RowCounter
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 7).FormulaLocal = FormulaFR
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 7) = SubstringID
            'PT
            FormulaPT = "= Substrings!H" & RowCounter
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 8).FormulaLocal = FormulaPT
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 8) = SubstringID
            'ES
            FormulaES = "= Substrings!I" & RowCounter
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 9).FormulaLocal = FormulaES
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 9) = SubstringID
            'DE
            FormulaDE = "= Substrings!J" & RowCounter
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 10).FormulaLocal = FormulaDE
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 10) = SubstringID
            'RO
            FormulaRO = "= Substrings!K" & RowCounter
            ActiveWorkbook.Sheets("CellReferenceParsing").Cells(i, 11).FormulaLocal = FormulaRO
            ActiveWorkbook.Sheets("UIDParsing").Cells(i, 11) = SubstringID
            'Incriminates RowCounter variable by 1
            RowCounter = RowCounter + 1
        End If
    End If
Next i
'Locks CellReferenceParsing sheet
ActiveWorkbook.Worksheets("CellReferenceParsing").Cells.Locked = True
ActiveWorkbook.Worksheets("CellReferenceParsing").Protect Password:="1", AllowFiltering:=True, AllowSorting:=True
'Adds filters and freezes first row on Substrings sheet
ActiveWorkbook.Worksheets("Substrings").Activate
ActiveWorkbook.Worksheets("Substrings").Rows("2:2").Select
ActiveWindow.FreezePanes = True
ActiveWorkbook.Worksheets("Substrings").Range("A1").AutoFilter
End Sub
