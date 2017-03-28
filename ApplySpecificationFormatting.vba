Sub ApplySpecificationFormatting()
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
            If LowerCaseText = "ñáîðî÷íûå åäèíèöû" Or LowerCaseText = "äîêóìåíòàöèÿ" Or LowerCaseText = "ñîñòàâ îáîðóäîâàíèÿ òåðìèíàëà" Or LowerCaseText = "ñòàíäàðòíûå èçäåëèÿ" _
            Or LowerCaseText = "ïðîãðàììíûå êîìïîíåíòû" Or LowerCaseText = "ïåðåìåííûå êîìïëåêòóþùèå" Or LowerCaseText = "ïåðåìåííûå äàííûå äëÿ èñïîëíåíèé" Or LowerCaseText = "âàðèàíòû èñïîëíåíèÿ ïàáê" _
            Or LowerCaseText = "äîêóìåíòàöèÿ èñòî÷íèêîâ ñîáûòèé" _
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
Next Row
End Sub
