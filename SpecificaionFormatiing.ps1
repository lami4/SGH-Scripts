Sub Test()
Dim Column As Range
Dim RowsCount As Integer
RowsCount = ActiveDocument.Tables(1).Rows.Count
MsgBox (RowsCount)
Set Column = ActiveDocument.Range(Start:=ActiveDocument.Tables(1).Cell(1, 1).Range.Start, _
End:=ActiveDocument.Tables(1).Cell(RowsCount, 1).Range.End)
End Sub

$application = New-Object -ComObject word.application
$application.Visible = $true
$document = $application.documents.open("C:\Users\Светлана\Downloads\Спецификация_final.docx")
$table = $document.Tables.Item(1)
$column = $document.Range([ref]$table.Cell(1, 1).Range.Start, [ref]$table.Cell(22, 1).Range.End)
$column.Select()
$column.Bold = $true
====================================VBA
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
'Main table. Columns 1-3.
For i = 1 To 3
ActiveDocument.Tables(1).Columns(i).Cells.VerticalAlignment = wdCellAlignVerticalCenter
ActiveDocument.Tables(1).Columns(i).Select
With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Bold = False
        .Italic = False
        .Underline = False
End With
With Selection.ParagraphFormat
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
Next i
'Main table. Column 6.
ActiveDocument.Tables(1).Columns(6).Cells.VerticalAlignment = wdCellAlignVerticalCenter
ActiveDocument.Tables(1).Columns(6).Select
With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Bold = False
        .Italic = False
        .Underline = False
End With
With Selection.ParagraphFormat
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
'Main table. Column 4.
ActiveDocument.Tables(1).Columns(4).Cells.VerticalAlignment = wdCellAlignVerticalCenter
ActiveDocument.Tables(1).Columns(4).Select
With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Bold = False
        .Italic = False
        .Underline = False
End With
With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
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
'Main table. Formatting Column 5.
ActiveDocument.Tables(1).Columns(5).Cells.VerticalAlignment = wdCellAlignVerticalCenter
ActiveDocument.Tables(1).Columns(5).Select
With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Bold = False
        .Italic = False
'        .Underline = False
End With
With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
'        .Alignment = wdAlignParagraphLeft
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
'Main table. Formatting section names in Column 5.
CellCounter = ActiveDocument.Tables(1).Columns(5).Cells.Count
For i = 1 To CellCounter
TextInCell = ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Text
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
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Size = 12
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Name = "Arial"
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Underline = True
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Bold = True
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Italic = False
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
End If
Next i
'Main table. Formatting Column 7.
ActiveDocument.Tables(1).Columns(7).Cells.VerticalAlignment = wdCellAlignVerticalCenter
ActiveDocument.Tables(1).Columns(7).Select
With Selection.Font
        .Name = "Arial"
        .Size = 6
        .Bold = False
        .Italic = False
        .Underline = False
End With
With Selection.ParagraphFormat
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
CellCounter = ActiveDocument.Tables(1).Columns(7).Cells.Count
For i = 1 To CellCounter
TextInCell = ActiveDocument.Tables(1).Columns(7).Cells(i).Range.Text
If Len(TextInCell) < 5 Then
ActiveDocument.Tables(1).Columns(7).Cells(i).Range.Font.Size = 8
End If
If Len(TextInCell) < 62 And TextInCell Like "*md5*" Then
ActiveDocument.Tables(1).Columns(7).Cells(i).Range.Font.Size = 6
End If
If Len(TextInCell) > 70 Then
ActiveDocument.Tables(1).Columns(7).Cells(i).Range.Font.Size = 4.5
ActiveDocument.Tables(1).Columns(7).Cells(i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
End If
Next i
'Removes linebreak before first character in a cell
For Each Column In ActiveDocument.Tables(1).Columns
CellCounter = Column.Cells.Count
For i = 1 To CellCounter
TextInCell = Column.Cells(i).Range.Text
TextInCell = Replace(TextInCell, ChrW(13) & ChrW(7), "")
LbPosition = InStr(TextInCell, ChrW(13))
If LbPosition = 1 Then
TextInCell = Replace(TextInCell, ChrW(13), "", 1, 1)
TextInCell = Trim(TextInCell)
Column.Cells(i).Range.Delete
Column.Cells(i).Range.Text = TextInCell
End If
Next i
Next Column
End Sub
=========================================================================================
