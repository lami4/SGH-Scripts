clear
#Создать экземпляр приложения MS Word
$word = New-Object -ComObject Word.Application
#Сделать вызванное приложение невидемым
$word.Visible = $true
#Создать документ MS Word
$document = $word.Documents.Add()

#НАСТРОЙКА ПОЛЕЙ ДОКУМЕНТА
#Левое поле (сантиметры)
$document.PageSetup.LeftMargin = $word.CentimetersToPoints(1)
#Правое поле (сантиметры)
$document.PageSetup.RightMargin = $word.CentimetersToPoints(1)
#Верхнее поле (сантиметры)
$document.PageSetup.TopMargin = $word.CentimetersToPoints(0.5)
#Нижнее поле (сантиметры)
$document.PageSetup.BottomMargin = $word.CentimetersToPoints(0.5)
#Верхний колонтитул
$document.PageSetup.HeaderDistance = $word.CentimetersToPoints(1)
#Нижний колонтитул
$document.PageSetup.FooterDistance = $word.CentimetersToPoints(1)

#НАСТРОЙКА ТАБЛИЦЫ
#Добавить таблицу
$document.Tables.Add($word.Selection.Range, 1, 1)
$table = $document.Tables.Item(1)
#Сделать границы таблицы видимыми
$table.Borders.Enable = $true
#Ширина таблицы (19,4 см)
$table.Columns.Item(1).Width = $word.CentimetersToPoints(19.4)
#Отсутуп от левого края в ячейках
$table.LeftPadding = $word.CentimetersToPoints(0.05)
#Отсутуп от правого края в ячейках
$table.RightPadding = $word.CentimetersToPoints(0.05)
#Установить вертикальное выравнивание по центру для всех ячеек
$table.Cell(1, 1).VerticalAlignment = 1
#Настройка шрифта в таблице
$table.Cell(1, 1).Range.Font.Name = "Arial"
$table.Cell(1, 1).Range.Font.Size = 10
#Интервал после (0 пт) для каждой ячейки таблицы
$table.Range.ParagraphFormat.SpaceAfter = 0
#Междустрочный интервал (одинарный) для каждой ячейки таблицы
$table.Range.ParagraphFormat.LineSpacingRule = 0

#ВЕРСТКА ТАБЛИЦЫ
#Разбить первую строку на 4 колонки
$table.Cell(1, 1).Split(1, 4)
$table.Cell(1, 1).Column.Width = $word.CentimetersToPoints(3)
$table.Cell(1, 2).Column.Width = $word.CentimetersToPoints(3)
$table.Cell(1, 3).Column.Width = $word.CentimetersToPoints(6.7)
$table.Cell(1, 4).Column.Width = $word.CentimetersToPoints(6.7)
$table.Rows.Add()
$table.Rows.Add()
$table.Cell(1, 1).Merge($table.Cell(2, 1))
$table.Cell(1, 2).Merge($table.Cell(2, 2))

#Добавить строку с датами и информации о количестве страниц
$table.Rows.Add()
$table.Cell(3, 4).Split(1, 3)
$table.Cell(3, 4).SetWidth($word.CentimetersToPoints(2.7), 1)
$table.Cell(3, 2).Merge($table.Cell(3, 4))
$document.Range([ref]$table.Cell(3, 1).Range.Start, [ref]$table.Cell(3, 4).Range.Start).Select()
$document.Application.Selection.InsertRowsBelow()
$Document.Application.Selection.Move()

#Вставить текст
$table.Cell(1, 2).Range.Text = "БТД"
$table.Cell(1, 3).Range.Text = "Извещение"
$table.Cell(1, 4).Range.Text = "Обозначение изменяемого документа"
$table.Cell(3, 1).Range.Text = "Дата выпуска"
$table.Cell(3, 2).Range.Text = "Срок внесения изменений"
$table.Cell(3, 3).Range.Text = "Лист"
$table.Cell(3, 4).Range.Text = "Листов"

#Добавить логотип компании и выровнять по центру
$table.Cell(1, 1).Range.InlineShapes.AddPicture("$PSScriptRoot\logo.jpg", $false, $true)
$table.Cell(1, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(1, 1).Range.ParagraphFormat.SpaceBefore = 1

#ФОРМАТИРОВАНИЕ ТАБЛИЦЫ ПОСЛЕ ВЕРСТКИ
#Высота строки (0,5 см) для всех ячеек
$table.Rows.Height = $word.CentimetersToPoints(0.5)
#Включить постоянную высоту для строк таблицы
$table.Rows.HeightRule = 2
$table.Cell(2, 3).Height = $word.CentimetersToPoints(1)
#Вставить надписть ГОСТ
$document.Range([ref]$table.Cell(1, 1).Range.Start, [ref]$table.Cell(1, 4).Range.Start).Select()
$document.Application.Selection.InsertRowsAbove()
$table.Cell(1, 1).Merge($table.Cell(1, 4))
$table.Cell(1, 1).Range.ParagraphFormat.Alignment = 2
$table.Cell(1, 1).Range.Text = "ГОСТ 2.503-90 Форма 1"
$table.Cell(1, 1).Borders.Item(-2).LineStyle = 0
$table.Cell(1, 1).Borders.Item(-1).LineStyle = 0
$table.Cell(1, 1).Borders.Item(-4).LineStyle = 0

#ДОБАВИТЬ НАДПИСИ
$document.PageSetup.DifferentFirstPageHeaderFooter = -1
$document.Sections.Item(1).Headers.Item(2).Shapes.AddShape(1, 10, 10, 200, 20).TextFrame.TextRange.Text = "Конфиденциально"
$shapeTop = $document.Sections.Item(1).Headers.Item(2).Shapes.Item(1)
$shapeTop.Height = $word.CentimetersToPoints(0.8)
$shapeTop.Width = $word.CentimetersToPoints(8.5)
$shapeTop.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$shapeTop.TextFrame.TextRange.Font.Size = 16
$shapeTop.TextFrame.TextRange.Font.Name = "Arial"
$shapeTop.TextFrame.TextRange.Font.Bold = $true
$shapeTop.TextFrame.TextRange.Font.ColorIndex = 1
$shapeTop.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
$shapeTop.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = 0
$shapeTop.RelativeHorizontalPosition = 5
$shapeTop.Left = $word.CentimetersToPoints(-8.2)
$shapeTop.RelativeVerticalPosition = 4
$shapeTop.Top = $word.CentimetersToPoints(0.4)
$shapeTop.Fill.Visible = 0
$shapeTop.Line.Weight = 1
$shapeTop.Line.Visible = 0
$shapeTop.TextFrame.MarginBottom = $word.CentimetersToPoints(0)
$shapeTop.TextFrame.MarginLeft = $word.CentimetersToPoints(0.1)
$shapeTop.TextFrame.MarginRight = $word.CentimetersToPoints(0.1)
$shapeTop.TextFrame.MarginTop = $word.CentimetersToPoints(0)
$footer = $document.Sections.Item(1).Footers.Item(2)
$footer.Range.Tables.Add($footer.Range, 1, 1)
$footer.Range.Tables.Item(1).Borders.Enable = $true
