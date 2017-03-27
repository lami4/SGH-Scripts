Sub TestTable()
izm = ActiveDocument.Tables(1).Cell(7, 3).Range.Text
izveshenie = ActiveDocument.Tables(1).Cell(7, 5).Range.Text
oboznachenie = ActiveDocument.Tables(1).Cell(6, 8).Range.Text
nazvanie = ActiveDocument.Tables(1).Cell(9, 7).Range.Text
izm = Mid(izm, 1, Len(izm) - 1)
izveshenie = Mid(izveshenie, 1, Len(izveshenie) - 1)
oboznachenie = Mid(oboznachenie, 1, Len(oboznachenie) - 1)
nazvanie = Mid(nazvanie, 1, Len(nazvanie) - 1)
MsgBox ("Изм.: " & izm & vbNewLine & "Номер извещения: " & izveshenie & vbNewLine & "Обозначение: " & oboznachenie & vbNewLine & vbNewLine & "Название: " & nazvanie)
End Sub
Sub ApplyTitleFormatting()
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.ParagraphFormat.SpaceAfter = 0
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.ParagraphFormat.SpaceBefore = 0
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Font.Name = "Arial"
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Font.Size = 10
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.ParagraphFormat.SpaceBefore = 0
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.ParagraphFormat.SpaceBefore = 0
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Font.Name = "Arial"
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Font.Size = 8
ActiveDocument.Sections(1).Footers(1).PageNumbers.NumberStyle = wdPageNumberStyleArabic
ActiveDocument.Sections(1).Footers(1).PageNumbers.StartingNumber = 1
Set part1 = ActiveDocument.Tables(1).Cell(2, 1).Range
part1.SetRange Start:=part1.Start, _
End:=ActiveDocument.Tables(1).Cell(9, 2).Range.End
With part1.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With part1.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
Set part2 = ActiveDocument.Tables(1).Cell(9, 3).Range
part2.SetRange Start:=part2.Start, _
End:=ActiveDocument.Tables(1).Cell(13, 6).Range.End
With part2.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With part2.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphLeft
End With
'Signature cells
signature1 = ActiveDocument.Tables(1).Cell(9, 5)
With signature1.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With signature1.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
signature2 = ActiveDocument.Tables(1).Cell(10, 5)
With signature1.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With signature2.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
signature3 = ActiveDocument.Tables(1).Cell(11, 5)
With signature1.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With signature3.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
signature4 = ActiveDocument.Tables(1).Cell(12, 5)
With signature1.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With signature4.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
signature5 = ActiveDocument.Tables(1).Cell(13, 5)
With signature1.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With signature5.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
'Date cells
date1 = ActiveDocument.Tables(1).Cell(9, 6)
With date1.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With date1.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
date2 = ActiveDocument.Tables(1).Cell(10, 6)
With date2.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With date2.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
date3 = ActiveDocument.Tables(1).Cell(11, 6)
With date3.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With date3.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
date4 = ActiveDocument.Tables(1).Cell(12, 6)
With date4.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With date4.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
date5 = ActiveDocument.Tables(1).Cell(13, 6)
With date5.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With date5.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
'Title
docTitle = ActiveDocument.Tables(1).Cell(9, 7)
With docTitle.Font
.Bold = False
.Name = "Arial"
.Size = 10
End With
With docTitle.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
'Name
docName = ActiveDocument.Tables(1).Cell(6, 8)
With docName.Font
.Bold = False
.Name = "Arial"
.Size = 14
End With
With docName.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
'Cells above logo
Set aboveLogoPart1 = ActiveDocument.Tables(1).Cell(9, 8).Range
aboveLogoPart1.SetRange Start:=aboveLogoPart1.Start, _
End:=ActiveDocument.Tables(1).Cell(9, 10).Range.End
With aboveLogoPart1.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With aboveLogoPart1.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
Set aboveLogoPart2 = ActiveDocument.Tables(1).Cell(10, 8).Range
aboveLogoPart2.SetRange Start:=aboveLogoPart2.Start, _
End:=ActiveDocument.Tables(1).Cell(10, 12).Range.End
With aboveLogoPart2.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With aboveLogoPart2.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
aboveLogoPart2.Fields.Update
'Logo
logo = ActiveDocument.Tables(1).Cell(11, 8)
With logo.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With logo.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
'Under title
Set underLogo = ActiveDocument.Tables(1).Cell(14, 1).Range
underLogo.SetRange Start:=underLogo.Start, _
End:=ActiveDocument.Tables(1).Cell(14, 3).Range.End
With underLogo.Font
.Bold = False
.Name = "Arial"
.Size = 8
End With
With underLogo.ParagraphFormat
.SpaceAfter = 0
.SpaceBefore = 0
.LeftIndent = CentimetersToPoints(0)
.RightIndent = CentimetersToPoints(0)
.Alignment = wdAlignParagraphCenter
End With
'Header
header = ActiveDocument.Tables(1).Cell(1, 2)
header.ParagraphFormat.SpaceAfter = 12
header.Font.Name = "Times New Roman"
header.Font.Bold = True
header.Font.Size = 20
header.ParagraphFormat.Alignment = wdAlignParagraphCenter
'List of approvals
approvals = ActiveDocument.Tables(1).Tables(1)
approvals.ParagraphFormat.SpaceAfter = 0
approvals.Font.Bold = False
approvals.Font.Name = "Times New Roman"
approvals.Font.Size = 12
approvals.ParagraphFormat.Alignment = wdAlignParagraphRight
ActiveDocument.Tables(1).Tables(1).LeftPadding = PixelsToPoints(7)
ActiveDocument.Tables(1).Tables(1).RightPadding = PixelsToPoints(7)
'Place captions
For Each Shape In ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes
Confidential = Shape.TextFrame.TextRange.Text
Confidential = Mid(Confidential, 1, Len(Confidential) - 1)
LowerConfidential = LCase(Confidential)
If Shape.Width = 240.9 And LowerConfidential = "конфиденциально" Then
Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
Shape.Left = CentimetersToPoints(-8.2)
Shape.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
Shape.Top = CentimetersToPoints(0.4)
End If
If Shape.Width = 240.9 And LowerConfidential = "коммерческая тайна" Then
Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
Shape.Left = CentimetersToPoints(-8.2)
Shape.RelativeVerticalPosition = wdRelativeVerticalPositionBottomMarginArea
Shape.Top = CentimetersToPoints(0)
End If
Next
End Sub
Sub ApplyCustomPageMargins()
With ActiveDocument.Sections(1).PageSetup
.TopMargin = CentimetersToPoints(1)
.BottomMargin = CentimetersToPoints(1)
.LeftMargin = CentimetersToPoints(0.7)
.RightMargin = CentimetersToPoints(0.9)
.FooterDistance = CentimetersToPoints(0.9)
.HeaderDistance = CentimetersToPoints(0.9)
.Gutter = CentimetersToPoints(0)
End With
End Sub
Sub PositionCapturesInDocument()
For Each Shape In ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes
ImageText = Shape.TextFrame.TextRange.Text
ImageText = Mid(ImageText, 1, Len(ImageText) - 1)
If ImageText = "Коммерческая тайна" And Not Shape.Width = 240.9 Then
Shape.Height = CentimetersToPoints(0.8)
Shape.Width = CentimetersToPoints(8.6)
Shape.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
Shape.TextFrame.TextRange.Font.ColorIndex = wdBlack
Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
Shape.Left = CentimetersToPoints(11.55)
Shape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
Shape.Top = CentimetersToPoints(28)
End If
If ImageText = "Конфиденциально" And Not Shape.Width = 240.9 Then
Shape.Height = CentimetersToPoints(0.8)
Shape.Width = CentimetersToPoints(8.6)
Shape.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
Shape.TextFrame.TextRange.Font.ColorIndex = wdBlack
Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
Shape.Left = CentimetersToPoints(11.55)
Shape.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
Shape.Top = CentimetersToPoints(0.7)
End If
Next
End Sub
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
'Main table. Formatting section names in Column 5.
CellCounter = ActiveDocument.Tables(1).Columns(5).Cells.Count
For i = 1 To CellCounter
TextInCell = ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Text
TextInCell = Replace(TextInCell, ChrW(13) & ChrW(7), "")
TextInCell = Replace(TextInCell, ChrW(13), "")
TextInCell = Replace(TextInCell, ".", "")
TextInCell = Trim(TextInCell)
LowerCaseText = LCase(TextInCell)
If LowerCaseText = "сборочные единицы" Or LowerCaseText = "документация" Or LowerCaseText = "состав оборудования терминала" Or LowerCaseText = "стандартные изделия" _
Or LowerCaseText = "ïðîãðàììíûå êîìïîíåíòû" Or LowerCaseText = "ïåðåìåííûå êîìïëåêòóþùèå" Or LowerCaseText = "ïåðåìåííûå äàííûå äëÿ èñïîëíåíèé" Or LowerCaseText = "âàðèàíòû èñïîëíåíèÿ ïàáê" _
Or LowerCaseText = "assembly units" Or LowerCaseText = "documentation" Or LowerCaseText = "terminal hardware specifications" Or LowerCaseText = "standard items" _
Or LowerCaseText = "software components" Or LowerCaseText = "variable items" Or LowerCaseText = "variable data for various assemblies" Or LowerCaseText = "list of bhss assemblies" Then
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Size = 12
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Name = "Arial"
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Underline = True
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Bold = True
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.Font.Italic = False
ActiveDocument.Tables(1).Columns(5).Cells(i).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
End If
Next i
End Sub
