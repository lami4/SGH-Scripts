Sub CreateWatermarks()
    Dim TextInWatermark As String
    For Each Section In ActiveDocument.Sections
        'Удаляет существующие фигуры с надписями (если они есть) в колонтитулах каждого раздела
        DeleteAlreadyExistingWatermarks SectionObject:=Section
    Next Section
    'Затем перебирает все разделы документа
    For Each Section In ActiveDocument.Sections
        'Настраивает нижний и верхний колонтитулы в текущем разделе
        SetupHeadersFooters SectionObject:=Section
        'Создает прямоугольную фигуру с надписью "Строго конфиденциально" или "Конфиденциально" в колонтитулах текущего раздела
        NewWatermark TextInside:="Strictly confidential", SectionObject:=Section
        'Создает прямоугольную фигуру с надписью "Коммерческая тайна" в колонтитулах текущего раздела
        NewWatermark TextInside:="Trade secret", SectionObject:=Section
        'Далее, перебирает все фигуры в колонтитулах текущего раздела
        For Each Shape In Section.Headers(wdHeaderFooterPrimary).Shapes
            'Забирает значение строки в нижнем регистре из текущей фигуры
            TextInWatermark = GetLowerCasedTextFromShape(Shape)
            'Устанавливает положение текущей фигуры в зависимости от значения ее строки
            'Если значение строки - "Коммерческая тайна", то фигура размещается внизу
            'Если значение строки - "Строго конфиденциально" или "Конфиденциально", то фигура размещается вверху
            LocateWatermarksOnTitlePage TextInWatermark:=TextInWatermark, ShapeObject:=Shape
        Next Shape
    Next Section
    'Вырезает первый раздел (=титульный лист)
    ActiveDocument.Sections(1).Range.Cut
    'Применяет новые координаты для фигур в оставшейся части документа (т.е. в теле документа)
    For Each Shape In ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes
        'Забирает значение строки в нижнем регистре из текущей фигуры
        TextInWatermark = GetLowerCasedTextFromShape(Shape)
        'Если значение строки - "Коммерческая тайна", то устанавливает для фигуры соответсвующие параметры
        If TextInWatermark = "trade secret" Then
            Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            Shape.Left = CentimetersToPoints(11.55)
            Shape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            Shape.Top = CentimetersToPoints(28)
            Shape.Height = CentimetersToPoints(0.8)
            Shape.Width = CentimetersToPoints(8.5)
        End If
        'Если значение строки - "Строго конфиденциально" или "Конфиденциально", то устанавливает для фигуры соответсвующие параметры
        If TextInWatermark = "confidential" Or TextInWatermark = "strictly confidential" Then
            Shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            Shape.Left = CentimetersToPoints(11.55)
            Shape.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
            Shape.Top = CentimetersToPoints(0.7)
            Shape.Height = CentimetersToPoints(0.8)
            Shape.Width = CentimetersToPoints(8.5)
        End If
    Next Shape
    'Вставляет назад ранее вырезанный первый раздел (=титульный лист)
    ActiveDocument.Sections.Add Range:=ActiveDocument.Sections(1).Range
    Set EmptySection = ActiveDocument.Sections(1).Range
    EmptySection.Collapse Direction:=wdCollapseStart
    EmptySection.Paste
    ActiveDocument.Sections(2).Range.Delete
End Sub
Function NewWatermark(TextInside, ByVal SectionObject As Object)
    'Данная функция создает фигуру с заданным текстом внутри
    SectionObject.Headers(wdHeaderFooterPrimary).Shapes.AddShape(msoShapeRectangle, 10, 10, 200, 20).TextFrame.TextRange.Text = TextInside
End Function
Function LocateWatermarksOnTitlePage(TextInWatermark, ByVal ShapeObject As Object)
    'Если значение строки - "Строго конфиденциально" или "Конфиденциально", то функция SetWatermarkProperties установит для фигуры соответсвующие параметры
    If TextInWatermark = "confidential" Or TextInWatermark = "strictly confidential" Then
        SetWatermarkProperties RelativeVerticalPosition:=4, TopCoordinate:=0.4, ShapeObject:=ShapeObject
    End If
    'Если значение строки - "Коммерческая тайна", то функция SetWatermarkProperties установит для фигуры соответсвующие параметры
    If TextInWatermark = "trade secret" Then
        SetWatermarkProperties RelativeVerticalPosition:=5, TopCoordinate:=0, ShapeObject:=ShapeObject
    End If
End Function
Function SetWatermarkProperties(RelativeVerticalPosition, TopCoordinate, ByVal ShapeObject As Object)
    'Данная функция устанавливает параметры фигуры
    ShapeObject.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
    ShapeObject.Left = CentimetersToPoints(-8.2)
    ShapeObject.RelativeVerticalPosition = RelativeVerticalPosition
    ShapeObject.Top = CentimetersToPoints(TopCoordinate)
    ShapeObject.Height = CentimetersToPoints(0.8)
    ShapeObject.Width = CentimetersToPoints(8.5)
    ShapeObject.TextFrame.TextRange.Font.Name = "Arial"
    ShapeObject.TextFrame.TextRange.Font.Size = 14
    ShapeObject.Fill.Visible = msoFalse
    ShapeObject.Line.Weight = 1
    ShapeObject.Line.Visible = msoFalse
    ShapeObject.TextFrame.MarginBottom = CentimetersToPoints(0)
    ShapeObject.TextFrame.MarginLeft = CentimetersToPoints(0)
    ShapeObject.TextFrame.MarginRight = CentimetersToPoints(0.1)
    ShapeObject.TextFrame.MarginTop = CentimetersToPoints(0.1)
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
Function DeleteAlreadyExistingWatermarks(ByVal SectionObject As Object)
    'Перебирает все фигуры в колонтитулах
    For Each Shape In SectionObject.Headers(wdHeaderFooterPrimary).Shapes
        'Забирает значение строки в нижнем регистре из текущей фигуры
        TextInWatermark = GetLowerCasedTextFromShape(Shape)
        'Если значение строки - "Строго конфиденциально", "Конфиденциально" или "Коммерческая тайна", удаляет текущую фигуру
        If TextInWatermark = "confidential" Or TextInWatermark = "trade secret" Or TextInWatermark = "strictly confidential" Then
            Shape.Delete
        End If
    Next Shape
End Function
Function GetLowerCasedTextFromShape(ByVal ShapeObject As Object)
    'Забирает значение строки в текущей фигуре
    TextInWatermark = ShapeObject.TextFrame.TextRange.Text
    'Удаляет последний символ в строке (знак "Конец ячейки"), так как он не нужен
    TextInWatermark = Mid(TextInWatermark, 1, Len(TextInWatermark) - 1)
    'Все символы в строке преобрауются в нижний регистр
    TextInWatermark = LCase(TextInWatermark)
    GetLowerCasedTextFromShape = TextInWatermark
End Function
