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
