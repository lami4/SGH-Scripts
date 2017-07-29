Sub Test()
Dim TextInWatermark As String
'Перебирает все разделы документа
For Each Section In ActiveDocument.Sections
    'Создает прямоугольную фигуру с надписью "Строго конфиденциально" или "Конфиденциально" в колонтитулах текущего раздела
    NewWatermark TextInside:="Strictly confidential", SectionObject:=Section
    'Создает прямоугольную фигуру с надписью "Коммерческая тайна" в колонтитулах текущего раздела
    NewWatermark TextInside:="Trade secret", SectionObject:=Section
    'Далее, перебирает все фигуры в колонтитулах текущего раздела
    For Each Shape In Section.Headers(wdHeaderFooterPrimary).Shapes
        'Забирает значение строки в текущей фигуре
        TextInWatermark = Shape.TextFrame.TextRange.Text
        TextInWatermark = Mid(TextInWatermark, 1, Len(TextInWatermark) - 1)
        'Все символы в строке преобразованы в нижний регистр
        TextInWatermark = LCase(TextInWatermark)
        'Устанавливает положение текущей фигуры в зависимости от значения ее строки
        'Если значение строки - "Коммерческая тайна", то фигура размещается внизу
        'Если значение строки - "Строго конфиденциально" или "Конфиденциально", то фигура размещается вверху
        LocateWatermark TextInWatermark:=TextInWatermark, ShapeObject:=Shape
    Next Shape
Next Section
End Sub
Function NewWatermark(TextInside, ByVal SectionObject As Object)
SectionObject.Headers(wdHeaderFooterPrimary).Shapes.AddShape(msoShapeRectangle, 10, 10, 200, 20).TextFrame.TextRange.Text = TextInside
End Function
Function LocateWatermark(TextInWatermark, ByVal ShapeObject As Object)
If TextInWatermark = "confidential" Or TextInWatermark = "strictly confidential" Then
SetWatermarkProperties RelativeVerticalPosition:=4, TopCoordinate:=0.4, ShapeObject:=ShapeObject
End If
If TextInWatermark = "trade secret" Then
SetWatermarkProperties RelativeVerticalPosition:=5, TopCoordinate:=0, ShapeObject:=ShapeObject
End If
End Function
Function SetWatermarkProperties(RelativeVerticalPosition, TopCoordinate, ByVal ShapeObject As Object)
ShapeObject.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
ShapeObject.Left = CentimetersToPoints(-8.2)
ShapeObject.RelativeVerticalPosition = RelativeVerticalPosition
ShapeObject.Top = CentimetersToPoints(TopCoordinate)
ShapeObject.Height = CentimetersToPoints(0.8)
ShapeObject.Width = CentimetersToPoints(8.5)
ShapeObject.TextFrame.TextRange.Font.Size = 14
ShapeObject.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
ShapeObject.TextFrame.TextRange.Font.ColorIndex = wdBlack
End Function
