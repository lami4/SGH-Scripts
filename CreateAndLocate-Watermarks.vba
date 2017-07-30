Sub CreateWatermarks()
Dim TextInWatermark As String
'Перебирает все разделы документа
For Each Section In ActiveDocument.Sections
    'Настраивает нижний и верхний колонтитулы в текущем разделе
    SetupHeadersFooters SectionObject:=Section
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
Function SetupHeadersFooters(ByVal SectionObject As Object)
    'Проверяет включен ли параметр "Особый колонтитул для первой страницы" в текущем разделе
    DifferentFirstPageHedaerFooter = SectionObject.PageSetup.DifferentFirstPageHeaderFooter
    'Проверяет включен ли параметр  "Разные колонтитулы для четных и нечетных страниц" в текущем разделе
    OddAndEvenPagesHeaderFooter = SectionObject.PageSetup.OddAndEvenPagesHeaderFooter
    'Если параметр "Особый колонтитул для первой страницы" включен (-1, т.е. True) в текущем разделе, то скрипт отключает его
    If DifferentFirstPageHedaerFooter = -1 Then
    SectionObject.PageSetup.DifferentFirstPageHeaderFooter = False
    End If
    'Если параметр "Особый колонтитул для первой страницы" включен (-1, т.е. True) в текущем разделе, то скрипт отключает его
    If OddAndEvenPagesHeaderFooter = -1 Then
    SectionObject.PageSetup.OddAndEvenPagesHeaderFooter = False
    End If
    'Отключает настройку "Как в предыдущем разделе" для верхнего колонтитула текущего раздела
    SectionObject.Headers(wdHeaderFooterPrimary).LinkToPrevious = False
    'Отключает настройку "Как в предыдущем разделе" для нижнего колонтитула текущего раздела
    SectionObject.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
End Function
