Sub LocateWatermarksInTranslatedDocumentBody()
For Each Shape In ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes
ImageText = Shape.TextFrame.TextRange.Text
ImageText = Mid(ImageText, 1, Len(ImageText) - 1)
LowerCaseText = LCase(ImageText)
If Shape.Width > 241.5 And LowerCaseText = "trade secret" Then
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
If Shape.Width > 241.5 And LowerCaseText = "confidential" Then
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
Next
End Sub
