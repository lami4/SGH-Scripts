Sub ApplyTitleFormattingInTranslatedDocument()
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
header.ParagraphFormat.Alignment = wdAlignParagraphCenter
'List of approvals
approvals = ActiveDocument.Tables(1).Tables(1)
approvals.ParagraphFormat.SpaceAfter = 0
approvals.Font.Bold = False
approvals.Font.Name = "Times New Roman"
approvals.ParagraphFormat.Alignment = wdAlignParagraphRight
ActiveDocument.Tables(1).Tables(1).LeftPadding = PixelsToPoints(7)
ActiveDocument.Tables(1).Tables(1).RightPadding = PixelsToPoints(7)
'Fix size of the cell with title
TextInCell = ActiveDocument.Tables(1).Cell(9, 7).Range.Text
Length = Len(TextInCell)
If Length <= 140 Then
ActiveDocument.Tables(1).Cell(9, 7).Range.Font.Size = 10
End If
If Length > 140 Then
ActiveDocument.Tables(1).Cell(9, 7).Range.Font.Size = 9
End If
If Length > 190 Then
ActiveDocument.Tables(1).Cell(9, 7).Range.Font.Size = 8
End If
If Length > 215 Then
ActiveDocument.Tables(1).Cell(9, 7).Range.Font.Size = 7.5
End If
End Sub
