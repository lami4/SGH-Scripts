Sub ApplyTitleFormatting()
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
'Under  title
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
header.Font.Name = "Times New Roman"
header.Font.Bold = True
header.ParagraphFormat.Alignment = wdAlignParagraphCenter
approvals = ActiveDocument.Tables(1).Tables(1)
approvals.Font.Bold = False
approvals.Font.Name = "Times New Roman"
approvals.ParagraphFormat.Alignment = wdAlignParagraphRight
ActiveDocument.Tables(1).Tables(1).LeftPadding = PixelsToPoints(7)
ActiveDocument.Tables(1).Tables(1).RightPadding = PixelsToPoints(7)
'Place captions
ActiveDocument.Sections(1).Headers(1).Shapes(1).RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
ActiveDocument.Sections(1).Headers(1).Shapes(1).Left = CentimetersToPoints(-8.2)
ActiveDocument.Sections(1).Headers(1).Shapes(1).RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
ActiveDocument.Sections(1).Headers(1).Shapes(1).Top = CentimetersToPoints(0.4)
ActiveDocument.Sections(1).Footers(1).Shapes(2).RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
ActiveDocument.Sections(1).Footers(1).Shapes(2).Left = CentimetersToPoints(-8.2)
ActiveDocument.Sections(1).Footers(1).Shapes(2).RelativeVerticalPosition = wdRelativeVerticalPositionBottomMarginArea
ActiveDocument.Sections(1).Footers(1).Shapes(2).Top = CentimetersToPoints(0)
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
Sub TestTable()
izm = ActiveDocument.Tables(1).Cell(7, 3).Range.Text
izveshenie = ActiveDocument.Tables(1).Cell(7, 5).Range.Text
oboznachenie = ActiveDocument.Tables(1).Cell(6, 8).Range.Text
nazvanie = ActiveDocument.Tables(1).Cell(9, 7).Range.Text
izm = Mid(izm, 1, Len(izm) - 1)
izveshenie = Mid(izveshenie, 1, Len(izveshenie) - 1)
oboznachenie = Mid(oboznachenie, 1, Len(oboznachenie) - 1)
nazvanie = Mid(nazvanie, 1, Len(nazvanie) - 1)
MsgBox ("Изм.: " & izm & vbNewLine & "Номер извещения: " & izveshenie & vbNewLine & "Обозначение: " & oboznachenie & vbNewLine & "Название: " & nazvanie)
End Sub


