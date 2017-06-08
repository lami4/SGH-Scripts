Sub LocateWatermarksInTranslatedDocumentBody()
'Applies the following coordinates and formatting to all shapes in headers and footers
For Each Shape In ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes
ImageText = Shape.TextFrame.TextRange.Text
ImageText = Mid(ImageText, 1, Len(ImageText) - 1)
LowerCaseText = LCase(ImageText)
If LowerCaseText = "confidential" Or LowerCaseText = "strictly confidential" Then
Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
Shape.Left = CentimetersToPoints(-8.2)
Shape.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
Shape.Top = CentimetersToPoints(0.4)
Shape.Height = CentimetersToPoints(0.8)
Shape.Width = CentimetersToPoints(8.5)
Shape.TextFrame.TextRange.Font.Size = 14
Shape.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
Shape.TextFrame.TextRange.Font.ColorIndex = wdBlack
End If
If LowerCaseText = "trade secret" Then
Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
Shape.Left = CentimetersToPoints(-8.2)
Shape.RelativeVerticalPosition = wdRelativeVerticalPositionBottomMarginArea
Shape.Top = CentimetersToPoints(0)
Shape.Height = CentimetersToPoints(0.8)
Shape.Width = CentimetersToPoints(8.5)
Shape.TextFrame.TextRange.Font.Size = 14
Shape.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
Shape.TextFrame.TextRange.Font.ColorIndex = wdBlack
End If
Next Shape
'Cuts the first section (section with the title)
ActiveDocument.Sections(1).Range.Cut
'Applies the following coordinates and formatting to the rest of shapes in headers and footers
For Each Shape In ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes
ImageText = Shape.TextFrame.TextRange.Text
ImageText = Mid(ImageText, 1, Len(ImageText) - 1)
LowerCaseText = LCase(ImageText)
If LowerCaseText = "trade secret" Then
Shape.TextFrame.TextRange.Font.Size = 14
Shape.Height = CentimetersToPoints(0.8)
Shape.Width = CentimetersToPoints(8.6)
Shape.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
Shape.TextFrame.TextRange.Font.ColorIndex = wdBlack
Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
Shape.Left = CentimetersToPoints(11.55)
Shape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
Shape.Top = CentimetersToPoints(28)
End If
If LowerCaseText = "confidential" Or LowerCaseText = "strictly confidential" Then
Shape.TextFrame.TextRange.Font.Size = 14
Shape.Height = CentimetersToPoints(0.8)
Shape.Width = CentimetersToPoints(8.6)
Shape.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
Shape.TextFrame.TextRange.Font.ColorIndex = wdBlack
Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
Shape.Left = CentimetersToPoints(11.55)
Shape.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
Shape.Top = CentimetersToPoints(0.7)
End If
Next Shape
'Pastes back the first section that was cut earlier
ActiveDocument.Sections.Add Range:=ActiveDocument.Sections(1).Range
Set Range2 = ActiveDocument.Sections(1).Range
Range2.Collapse Direction:=wdCollapseStart
Range2.Paste
ActiveDocument.Sections(2).Range.Delete
End Sub
Sub ApplyFormattingInSpecification()
'Table in header
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
With ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Range.Font
        .Name = "Arial"
        .Size = 12
        .Bold = True
        .Italic = False
End With
With ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Range.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
Dim TextInCell As String
TextInCell = ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(2, 2).Range.Text
StringLength = Len(TextInCell)
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(2, 2).HeightRule = wdRowHeightExactly
If StringLength <= 34 Then
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(2, 2).Range.Font.Size = 10
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
Else
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(2, 2).Range.Font.Size = 8
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
End If
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(3, 1).Range.Orientation = wdTextOrientationUpward
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(3, 2).Range.Orientation = wdTextOrientationUpward
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(3, 3).Range.Orientation = wdTextOrientationUpward
ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(3, 6).Range.Orientation = wdTextOrientationUpward
'Table in footer
Dim TableInFooter As Range
Set TableInFooter = ActiveDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Range
TableInFooter.Cells.VerticalAlignment = wdCellAlignVerticalCenter
With TableInFooter.Font
        .Name = "Arial"
        .Size = 8
        .Bold = False
        .Italic = False
End With
With TableInFooter.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
Dim NamesAndFields As Range
Set NamesAndFields = ActiveDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(3, 1).Range
NamesAndFields.SetRange Start:=NamesAndFields.Start, End:=ActiveDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(8, 2).Range.End
NamesAndFields.Select
Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
ActiveDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(1, 6).Range.Font.Bold = True
ActiveDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(1, 6).Range.Font.Size = 14
ActiveDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(4, 5).Range.Font.Size = 10
TableCounter = ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables.Count
'Small table in footer (if any exists)
If TableCounter > 0 Then
Dim TableInFooterSmall As Range
Set TableInFooterSmall = ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Range
TableInFooterSmall.Cells.VerticalAlignment = wdCellAlignVerticalCenter
With TableInFooterSmall.Font
        .Name = "Arial"
        .Size = 8
        .Bold = False
        .Italic = False
End With
With TableInFooterSmall.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(1, 6).Range.Font.Bold = True
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(1, 6).Range.Font.Size = 14
End If
'Main table. Applying formatting.
'Column 7. Fixes the wrong numbering and ofrmatting of the numbered lists
Dim ListCounter As Integer
For Each RowItem In ActiveDocument.Tables(1).Rows
    ListCounter = 1
    If RowItem.Cells(7).Range.ListParagraphs.Count <> 0 Then
        RowItem.Cells(7).Range.ListFormat.RemoveNumbers
        For Each Paragraph In RowItem.Cells(7).Range.Paragraphs
            Paragraph.Range.InsertBefore (ListCounter & ". ")
            ListCounter = ListCounter + 1
        Next Paragraph
    End If
Next RowItem
'Deletes empty rows
For Each Row In ActiveDocument.Tables(1).Rows
TextInRow = Row.Range.Text
TextInRow = Replace(TextInRow, ChrW(13) & ChrW(7), "")
TextInRow = Replace(TextInRow, ChrW(13), "")
TextInRow = Replace(TextInRow, " ", "")
TextLength = Len(TextInRow)
If TextLength = 0 Then
Row.Delete
End If
Next Row
'Sets row height
ActiveDocument.Tables(1).Rows.Height = CentimetersToPoints(1)
ActiveDocument.Tables(1).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
For Each Row In ActiveDocument.Tables(1).Rows
    'if row contains more or less than 7 cells
    If Row.Cells.Count <> 7 Then
    For Each Cell In Row.Cells
        With Cell.Range.Font
            .Name = "Arial"
            .Size = 8
        End With
        Cell.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Next Cell
    Else
    'if row contains 7 cells
    For Each Cell In Row.Cells
'Main table. Columns 1, 2, 3, 4 and 6.
    If Cell.ColumnIndex <= 4 Or Cell.ColumnIndex = 6 Or Cell.ColumnIndex = 7 Then
        With Cell.Range.Font
            .Name = "Arial"
            .Size = 8
            .Bold = False
            .Italic = False
            .Underline = False
        End With
        With Cell.Range.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphCenter
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
        End With
    End If
'Main table. Apply special formatting for cells in column 4
    If Cell.ColumnIndex = 4 Then
        Cell.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    End If
'Main Table. Apply special formatting for cells in column 5
    If Cell.ColumnIndex = 5 Then
    'Make text left aligned if text in 4th column of the same row contains 'pabk'
        ColumnIndex = Cell.ColumnIndex
        ColumnIndex = ColumnIndex - 1
        TextInCell = Row.Cells(ColumnIndex)
        LowerCaseText = LCase(TextInCell)
        If LowerCaseText Like "*pabk*gl*" Then
        Cell.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        With Cell.Range.Font
            .Name = "Arial"
            .Size = 8
            .Bold = False
            .Italic = False
            .Underline = False
        End With
    'Make text left aligned if text in 4th column of the same row contains 'pabk' (above)
        End If
        With Cell.Range.Font
            .Name = "Arial"
            .Size = 8
            .Bold = False
            .Italic = False
'           .Underline = False
        End With
        With Cell.Range.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
'           .Alignment = wdAlignParagraphLeft
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
        End With
'Formatting section names in Column 5.
            TextInCell = Cell.Range.Text
            TextInCell = Replace(TextInCell, ChrW(13) & ChrW(7), "")
            TextInCell = Replace(TextInCell, ChrW(13), "")
            TextInCell = Replace(TextInCell, ".", "")
            TextInCell = Trim(TextInCell)
            LowerCaseText = LCase(TextInCell)
            If LowerCaseText = "сборочные единицы" Or LowerCaseText = "документация" Or LowerCaseText = "состав оборудования терминала" Or LowerCaseText = "стандартные изделия" _
            Or LowerCaseText = "программные компоненты" Or LowerCaseText = "переменные комплектующие" Or LowerCaseText = "переменные данные для исполнений" Or LowerCaseText = "варианты исполнения пабк" _
            Or LowerCaseText = "документация источников событий" _
            Or LowerCaseText = "assembly units" Or LowerCaseText = "documentation" Or LowerCaseText = "terminal hardware specifications" Or LowerCaseText = "standard items" _
            Or LowerCaseText = "software components" Or LowerCaseText = "variable items" Or LowerCaseText = "variable data for various assemblies" Or LowerCaseText = "list of bhss assemblies" _
            Or LowerCaseText = "event source documentation" Then
            Cell.Range.Font.Size = 12
            Cell.Range.Font.Name = "Arial"
            Cell.Range.Font.Underline = True
            Cell.Range.Font.Bold = True
            Cell.Range.Font.Italic = False
            Cell.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            End If
    End If
'Main Table. Apply special formatting for cells in column 7
    If Cell.ColumnIndex = 7 Then
        Cell.Range.Font.Size = 6
        TextInCell = Cell.Range.Text
        LowerCaseText = LCase(TextInCell)
        If Len(LowerCaseText) < 5 Then
        Cell.Range.Font.Size = 8
        End If
        If Len(LowerCaseText) < 62 And LowerCaseText Like "*md5*" Then
        Cell.Range.Font.Size = 6
        End If
        If Len(LowerCaseText) > 70 Then
        Cell.Range.Font.Size = 4.5
        Cell.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End If
    End If
    For i = 1 To 3
'Main Table. Removes linebreak before first character in a cell
        TextInCell = Cell.Range.Text
        TextInCell = Replace(TextInCell, ChrW(13) & ChrW(7), "")
        LbPosition = InStr(TextInCell, ChrW(13))
        If LbPosition = 1 Then
            TextInCell = Replace(TextInCell, ChrW(13), "", 1, 1)
            TextInCell = Trim(TextInCell)
            Cell.Range.Delete
            Cell.Range.Text = TextInCell
        End If
'Main Table. Removes linebreak after last character in a cell
        TextInCell = Cell.Range.Text
        TextInCell = Replace(TextInCell, ChrW(13) & ChrW(7), "")
        StringLength = Len(TextInCell)
        If Right$(TextInCell, 1) = ChrW(13) Then
            TextInCell = Left(TextInCell, Len(TextInCell) - 1)
            TextInCell = Trim(TextInCell)
            Cell.Range.Delete
            Cell.Range.Text = TextInCell
        End If
'Main Table. Removes space before first character in a cell
        TextInCell = Cell.Range.Text
        TextInCell = Replace(TextInCell, ChrW(13) & ChrW(7), "")
        LbPosition = InStr(TextInCell, " ")
        If LbPosition = 1 Then
            TextInCell = Replace(TextInCell, " ", "", 1, 1)
            TextInCell = Trim(TextInCell)
            Cell.Range.Delete
            Cell.Range.Text = TextInCell
        End If
'Main Table. Removes space after last character in a cell
        TextInCell = Cell.Range.Text
        TextInCell = Replace(TextInCell, ChrW(13) & ChrW(7), "")
        StringLength = Len(TextInCell)
        If Right$(TextInCell, 1) = " " Then
        TextInCell = Trim(TextInCell)
        Cell.Range.Delete
        Cell.Range.Text = TextInCell
        End If
    Next i
    Next Cell
    End If
Next Row
End Sub
