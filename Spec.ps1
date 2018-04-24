$script:counter = 1

Function Apply-FormattingInListTable($TableObject, $WordApp) {
#Сделать границы таблицы видимыми
$TableObject.Borders.Enable = $true
#Ширина таблицы (19,4 см)
$TableObject.Columns.Item(1).Width = $WordApp.CentimetersToPoints(18)
#Отсутуп от левого края в ячейках
$TableObject.LeftPadding = $WordApp.CentimetersToPoints(0.1)
#Отсутуп от правого края в ячейках
$TableObject.RightPadding = $WordApp.CentimetersToPoints(0.1)
#Установить вертикальное выравнивание по центру для всех ячеек
$TableObject.Cell(1, 1).VerticalAlignment = 1
#Настройка шрифта в таблице
$TableObject.Cell(1, 1).Range.Font.Name = "Arial"
$TableObject.Cell(1, 1).Range.Font.Size = 9
#Интервал после (0 пт) для каждой ячейки таблицы
$TableObject.Range.ParagraphFormat.SpaceAfter = 0
#Междустрочный интервал (одинарный) для каждой ячейки таблицы
$TableObject.Range.ParagraphFormat.LineSpacingRule = 0
#Установить высоту на минимум
$TableObject.Cell(15, 1).HeightRule = 1
$TableObject.Rows.Height = $word.CentimetersToPoints(0.5)
#Разбить таблицу на 4 столбца и установить их ширину
$TableObject.Cell(1, 1).Split(1, 4)
$TableObject.Cell(1, 1).Column.Width = $WordApp.CentimetersToPoints(1.5)
$TableObject.Cell(1, 2).Column.Width = $WordApp.CentimetersToPoints(1.5)
$TableObject.Cell(1, 3).Column.Width = $WordApp.CentimetersToPoints(7.5)
$TableObject.Cell(1, 4).Column.Width = $WordApp.CentimetersToPoints(7.5)
#Добавить надписи в таблицу
$TableObject.Cell(1, 1).Range.Text = "Поз."
$TableObject.Cell(1, 2).Range.Text = "Изм."
$TableObject.Cell(1, 3).Range.Text = "Обозначение"
$TableObject.Cell(1, 4).Range.Text = "Примечание"
#Установить выравниевание по центру для заголовка
$TableObject.Rows.Item(1).Range.ParagraphFormat.Alignment = 1
#Выровнять таблицу по центру в документу
$TableObject.Rows.Alignment = 1
#Добавить строку для входных данных
$TableObject.Rows.Add()
#Запретить переносить ячейку на селдующую страницу
$TableObject.Rows.Item(2).AllowBreakAcrossPages = $false
#Отформатировать строку для входных данных
$TableObject.Cell(2, 3).Range.ParagraphFormat.Alignment = 0
$TableObject.Cell(2, 2).LeftPadding = $word.CentimetersToPoints(0.1)
$TableObject.Cell(2, 2).RightPadding = $word.CentimetersToPoints(0.1)
}

Function Add-TestData($TableObject) {
    for ($t = 2; $t -lt 40; $t++) {
    $TableObject.Cell($t, 1).Range.Text = $script:counter
    $script:counter += 1
    $TableObject.Cell($t, 2).Range.Text = "-"
    $TableObject.Cell($t, 3).Range.Text = "ABCDFE-RU-RU-00.00.00.tEST.01.00_00.11"
    $TableObject.Cell($t, 4).Range.Text = "81cf2f9f23fd597f2e278e56718c3831"
    $TableObject.Rows.Add()
    }
}

clear
#Создать экземпляр приложения MS Word
$word = New-Object -ComObject Word.Application
#Создать документ MS Word
$document = $word.Documents.Add()
#Сделать вызванное приложение невидемым
$word.Visible = $true
Start-Sleep -Seconds 5
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
$table.Cell(1, 1).Range.Font.Size = 9
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
$document.Range([ref]$table.Cell(3, 1).Range.Start, [ref]$table.Cell(3, 4).Range.Start).Select()
$document.Application.Selection.InsertRowsBelow()
$document.Application.Selection.InsertRowsBelow()
$document.Application.Selection.InsertRowsBelow()
$Document.Application.Selection.Move()
$table.Cell(3, 2).Merge($table.Cell(3, 4))
$table.Cell(4, 2).Merge($table.Cell(4, 4))
$table.Cell(5, 3).Merge($table.Cell(5, 4))
$table.Cell(6, 3).Merge($table.Cell(6, 4))
$table.Cell(5, 1).Merge($table.Cell(5, 2))
$table.Cell(6, 1).Merge($table.Cell(6, 2))
$table.Cell(5, 3).Merge($table.Cell(5, 4))
$table.Cell(6, 3).Merge($table.Cell(6, 4))
$table.Cell(5, 1).Merge($table.Cell(6, 1))
$table.Cell(5, 2).Merge($table.Cell(6, 2))
$table.Cell(7, 1).Merge($table.Cell(7, 2))
$table.Cell(7, 2).Merge($table.Cell(7, 3))
$table.Rows.Add()
$table.Rows.Add()
$table.Rows.Add()
$table.Rows.Add()
$table.Rows.Add()
$table.Cell(12, 1).SetWidth($word.CentimetersToPoints(1.5), 1)
$table.Rows.Add()
$table.Rows.Add()
$table.Cell(12, 2).Merge($table.Cell(13, 2))
$table.Cell(14, 1).Merge($table.Cell(14, 2))

for ($i = 0; $i -lt 29; $i++) {
$table.Rows.Add()
}

#Вставить текст
$table.Cell(1, 2).Range.Text = "БТД"
$table.Cell(1, 3).Range.Text = "Извещение"
$table.Cell(1, 4).Range.Text = "Обозначение изменяемого документа"
$table.Cell(3, 1).Range.Text = "Дата выпуска"
$table.Cell(3, 2).Range.Text = "Срок внесения изменений"
$table.Cell(3, 3).Range.Text = "Лист"
$table.Cell(3, 4).Range.Text = "Листов"
$table.Cell(5, 1).Range.Text = "Причина"
$table.Cell(5, 3).Range.Text = "Код"
$table.Cell(7, 1).Range.Text = "Указание о заделе"
$table.Cell(8, 1).Range.Text = "Указание о внедрении"
$table.Cell(9, 1).Range.Text = "Применяемость"
$table.Cell(10, 1).Range.Text = "Разослать"
$table.Cell(11, 1).Range.Text = "Приложение"
$table.Cell(12, 1).Range.Text = "Изм."
$table.Cell(13, 1).Range.Text = "-"
$table.Cell(12, 2).Range.Text = "Содержание изменения"
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
$table.Cell(7, 1).Height = $word.CentimetersToPoints(1)
$table.Cell(8, 1).Height = $word.CentimetersToPoints(1)
$table.Cell(9, 1).Height = $word.CentimetersToPoints(1)
$table.Cell(10, 1).Height = $word.CentimetersToPoints(1)
$table.Cell(11, 1).Height = $word.CentimetersToPoints(1)
#Настроить выравнивание в ячейках
$table.Cell(1, 2).Range.ParagraphFormat.Alignment = 1
$table.Cell(5, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(7, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(8, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(9, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(10, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(11, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(12, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(3, 3).Range.ParagraphFormat.Alignment = 1
$table.Cell(3, 4).Range.ParagraphFormat.Alignment = 1
$table.Cell(5, 3).Range.ParagraphFormat.Alignment = 1
$table.Cell(13, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(12, 2).Range.ParagraphFormat.Alignment = 1
#Вставить надписть ГОСТ
$document.Range([ref]$table.Cell(1, 1).Range.Start, [ref]$table.Cell(1, 4).Range.Start).Select()
$document.Application.Selection.InsertRowsAbove()
$table.Cell(1, 1).Merge($table.Cell(1, 4))
$table.Cell(1, 1).Range.ParagraphFormat.Alignment = 2
$table.Cell(1, 1).Range.Text = "ГОСТ 2.503-90 Форма 1"
$table.Cell(1, 1).Borders.Item(-2).LineStyle = 0
$table.Cell(1, 1).Borders.Item(-1).LineStyle = 0
$table.Cell(1, 1).Borders.Item(-4).LineStyle = 0

#ДОБАВИТЬ НАДПИСЬ В ВЕРХНИЙ КОЛОНТИТУЛ НА ПЕРВОЙ СТРАНИЦЕ
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
$shapeTop.RelativeHorizontalPosition = 1
$shapeTop.Left = $word.CentimetersToPoints(11.8)
$shapeTop.RelativeVerticalPosition = 1
$shapeTop.Top = $word.CentimetersToPoints(0.4)
$shapeTop.Fill.Visible = 0
$shapeTop.Line.Weight = 1
$shapeTop.Line.Visible = 0
$shapeTop.TextFrame.MarginBottom = $word.CentimetersToPoints(0)
$shapeTop.TextFrame.MarginLeft = $word.CentimetersToPoints(0.1)
$shapeTop.TextFrame.MarginRight = $word.CentimetersToPoints(0.1)
$shapeTop.TextFrame.MarginTop = $word.CentimetersToPoints(0)

#НИЖНИЙ КОЛОНТИТУЛ НА ПЕРВОЙ СТРАНИЦЕ
$footer = $document.Sections.Item(1).Footers.Item(2)
$footer.Range.Tables.Add($footer.Range, 1, 1)
$footerTable = $footer.Range.Tables.Item(1)
$footerTable.Borders.Enable = $true
#Ширина таблицы (19,4 см)
$footerTable.Columns.Item(1).Width = $word.CentimetersToPoints(19.4)
#Отсутуп от левого края в ячейках
$footerTable.LeftPadding = $word.CentimetersToPoints(0.05)
#Отсутуп от правого края в ячейках
$footerTable.RightPadding = $word.CentimetersToPoints(0.05)
#Установить вертикальное выравнивание по центру для всех ячеек
$footerTable.Cell(1, 1).VerticalAlignment = 1
#Настройка шрифта в таблице
$footerTable.Cell(1, 1).Range.Font.Name = "Arial"
$footerTable.Cell(1, 1).Range.Font.Size = 9
#Интервал после (0 пт) для каждой ячейки таблицы
$footerTable.Range.ParagraphFormat.SpaceAfter = 0
#Междустрочный интервал (одинарный) для каждой ячейки таблицы
$footerTable.Range.ParagraphFormat.LineSpacingRule = 0
#Разбить таблицу на восемь колонок и задать их ширину
$footerTable.Cell(1, 1).Split(1, 8)
$footerTable.Cell(1, 1).Column.Width = $word.CentimetersToPoints(2)
$footerTable.Cell(1, 2).Column.Width = $word.CentimetersToPoints(3.5)
$footerTable.Cell(1, 3).Column.Width = $word.CentimetersToPoints(2.2)
$footerTable.Cell(1, 4).Column.Width = $word.CentimetersToPoints(2)
$footerTable.Cell(1, 5).Column.Width = $word.CentimetersToPoints(2)
$footerTable.Cell(1, 6).Column.Width = $word.CentimetersToPoints(3.5)
$footerTable.Cell(1, 7).Column.Width = $word.CentimetersToPoints(2.2)
$footerTable.Cell(1, 8).Column.Width = $word.CentimetersToPoints(2)
#Добавить еще четырке строки в таблицу
$footerTable.Rows.Add()
$footerTable.Rows.Add()
$footerTable.Rows.Add()
$footerTable.Rows.Add()
#Добавить надписи в таблицу
$footerTable.Cell(1, 2).Range.Text = "Фамилия"
$footerTable.Cell(1, 3).Range.Text = "Подпись"
$footerTable.Cell(1, 4).Range.Text = "Дата"
$footerTable.Cell(1, 6).Range.Text = "Фамилия"
$footerTable.Cell(1, 7).Range.Text = "Подпись"
$footerTable.Cell(1, 8).Range.Text = "Дата"
$footerTable.Cell(2, 1).Range.Text = "Составил"
$footerTable.Cell(3, 1).Range.Text = "Проверил"
$footerTable.Cell(4, 1).Range.Text = "Т. контр."
$footerTable.Cell(5, 1).Range.Text = "Изменение внес"
$footerTable.Cell(2, 5).Range.Text = "Н. контр."
$footerTable.Cell(4, 5).Range.Text = "Утвердил"
$footerTable.Rows.Item(1).Range.ParagraphFormat.Alignment = 1
#Объединить строку "Изменения внес"
$footerTable.Cell(5, 1).Merge($footerTable.Cell(5, 8))
#Высота строки (0,5 см) для всех ячеек
$footerTable.Rows.Height = $word.CentimetersToPoints(0.5)
#Включить постоянную высоту для строк таблицы
$footerTable.Rows.HeightRule = 2

#ДОБАВИТЬ НАДПИСЬ В НИЖНИЙ КОЛОНТИТУЛ НА ПЕРВОЙ СТРАНИЦЕ
$document.Sections.Item(1).Headers.Item(2).Shapes.AddShape(1, 10, 10, 200, 20).TextFrame.TextRange.Text = "Коммерческая тайна"
$shapeBot = $document.Sections.Item(1).Headers.Item(2).Shapes.Item(2)
$shapeBot.Height = $word.CentimetersToPoints(0.8)
$shapeBot.Width = $word.CentimetersToPoints(8.5)
$shapeBot.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$shapeBot.TextFrame.TextRange.Font.Size = 16
$shapeBot.TextFrame.TextRange.Font.Name = "Arial"
$shapeBot.TextFrame.TextRange.Font.Bold = $true
$shapeBot.TextFrame.TextRange.Font.ColorIndex = 1
$shapeBot.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
$shapeBot.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = 0
$shapeBot.RelativeHorizontalPosition = 1
$shapeBot.Left = $word.CentimetersToPoints(11.8)
$shapeBot.RelativeVerticalPosition = 1
$shapeBot.Top = $word.CentimetersToPoints(28.4)
$shapeBot.Fill.Visible = 0
$shapeBot.Line.Weight = 1
$shapeBot.Line.Visible = 0
$shapeBot.TextFrame.MarginBottom = $word.CentimetersToPoints(0)
$shapeBot.TextFrame.MarginLeft = $word.CentimetersToPoints(0.1)
$shapeBot.TextFrame.MarginRight = $word.CentimetersToPoints(0.1)
$shapeBot.TextFrame.MarginTop = $word.CentimetersToPoints(0)

#ВТОРАЯ СТРАНИЦА
#Верхний колонтитул

#Надпись в верхнем колонтитуле
$document.Sections.Item(1).Headers.Item(1).Shapes.AddShape(1, 10, 10, 200, 20).TextFrame.TextRange.Text = "Конфиденциально"
$shapeTopPageTwo = $document.Sections.Item(1).Headers.Item(1).Shapes.Item(3)
$shapeTopPageTwo.Height = $word.CentimetersToPoints(0.8)
$shapeTopPageTwo.Width = $word.CentimetersToPoints(8.5)
$shapeTopPageTwo.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$shapeTopPageTwo.TextFrame.TextRange.Font.Size = 16
$shapeTopPageTwo.TextFrame.TextRange.Font.Name = "Arial"
$shapeTopPageTwo.TextFrame.TextRange.Font.Bold = $true
$shapeTopPageTwo.TextFrame.TextRange.Font.ColorIndex = 1
$shapeTopPageTwo.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
$shapeTopPageTwo.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = 0
$shapeTopPageTwo.RelativeHorizontalPosition = 1
$shapeTopPageTwo.RelativeVerticalPosition = 1
$shapeTopPageTwo.Left = $word.CentimetersToPoints(11.8)
$shapeTopPageTwo.Top = $word.CentimetersToPoints(0.4)
$shapeTopPageTwo.Fill.Visible = 0
$shapeTopPageTwo.Line.Weight = 1
$shapeTopPageTwo.Line.Visible = 0
$shapeTopPageTwo.TextFrame.MarginBottom = $word.CentimetersToPoints(0)
$shapeTopPageTwo.TextFrame.MarginLeft = $word.CentimetersToPoints(0.1)
$shapeTopPageTwo.TextFrame.MarginRight = $word.CentimetersToPoints(0.1)
$shapeTopPageTwo.TextFrame.MarginTop = $word.CentimetersToPoints(0)

#Таблица в верхнем колонтитуле
$header = $document.Sections.Item(1).Headers.Item(1).Range
$header.Collapse(0)
$document.Sections.Item(1).Headers.Item(1).Range.Tables.Add($header, 1, 1)
$headerTable = $document.Sections.Item(1).Headers.Item(1).Range.Tables.Item(1)
$headerTable.Borders.Enable = $true
#Ширина таблицы (19,4 см)
$headerTable.Columns.Item(1).Width = $word.CentimetersToPoints(19.4)
#Отсутуп от левого края в ячейках
$headerTable.LeftPadding = $word.CentimetersToPoints(0.05)
#Отсутуп от правого края в ячейках
$headerTable.RightPadding = $word.CentimetersToPoints(0.05)
#Установить вертикальное выравнивание по центру для всех ячеек
$headerTable.Cell(1, 1).VerticalAlignment = 1
#Настройка шрифта в таблице
$headerTable.Cell(1, 1).Range.Font.Name = "Arial"
$headerTable.Cell(1, 1).Range.Font.Size = 9
#Интервал после (0 пт) для каждой ячейки таблицы
$headerTable.Range.ParagraphFormat.SpaceAfter = 0
#Междустрочный интервал (одинарный) для каждой ячейки таблицы
$headerTable.Range.ParagraphFormat.LineSpacingRule = 0
#Вставить надписть ГОСТ
$headerTable.Cell(1, 1).Range.Text = "ГОСТ 2.503-90 Форма 1"
$headerTable.Rows.Add()
$headerTable.Cell(1, 1).Range.ParagraphFormat.Alignment = 2
$headerTable.Cell(1, 1).Borders.Item(-2).LineStyle = 0
$headerTable.Cell(1, 1).Borders.Item(-1).LineStyle = 0
$headerTable.Cell(1, 1).Borders.Item(-4).LineStyle = 0
#Собрать сотальную часть
$headerTable.Cell(2, 1).Split(1, 4)
$headerTable.Cell(2, 1).SetWidth($word.CentimetersToPoints(2.5), 1)
$headerTable.Cell(2, 2).SetWidth($word.CentimetersToPoints(13.9), 1)
$headerTable.Rows.Add()
$headerTable.Cell(4, 2).Merge($headerTable.Cell(4, 4))
$headerTable.Cell(4, 1).SetWidth($word.CentimetersToPoints(1.5), 1)
$headerTable.Rows.Add()
$headerTable.Cell(3, 2).Merge($headerTable.Cell(4, 2))
#Высота строки (0,5 см) для всех ячеек
$headerTable.Rows.Height = $word.CentimetersToPoints(0.5)
#Настройка выравнивания в ячейках
$headerTable.Cell(2, 1).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(2, 3).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(2, 4).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(3, 1).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(4, 1).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(3, 2).Range.ParagraphFormat.Alignment = 1
#Включить постоянную высоту для строк таблицы
$headerTable.Rows.HeightRule = 2
#Добавить текст в таблице верхнего колонтитула второй страницы
$headerTable.Cell(2, 1).Range.Text = "Извещение"
$headerTable.Cell(2, 3).Range.Text = "Лист"
$headerTable.Cell(3, 1).Range.Text = "Изм."
$headerTable.Cell(4, 1).Range.Text = "-"
$headerTable.Cell(3, 2).Range.Text = "Содержание изменения"
#Надпись в нижнем колонтитуле
$document.Sections.Item(1).Headers.Item(1).Shapes.AddShape(1, 10, 10, 200, 20).TextFrame.TextRange.Text = "Коммерческая тайна"
$shapeBotPageTwo = $document.Sections.Item(1).Headers.Item(1).Shapes.Item(4)
$shapeBotPageTwo.Height = $word.CentimetersToPoints(0.8)
$shapeBotPageTwo.Width = $word.CentimetersToPoints(8.5)
$shapeBotPageTwo.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$shapeBotPageTwo.TextFrame.TextRange.Font.Size = 16
$shapeBotPageTwo.TextFrame.TextRange.Font.Name = "Arial"
$shapeBotPageTwo.TextFrame.TextRange.Font.Bold = $true
$shapeBotPageTwo.TextFrame.TextRange.Font.ColorIndex = 1
$shapeBotPageTwo.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
$shapeBotPageTwo.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = 0
$shapeBotPageTwo.RelativeHorizontalPosition = 1
$shapeBotPageTwo.Left = $word.CentimetersToPoints(11.8)
$shapeBotPageTwo.RelativeVerticalPosition = 1
$shapeBotPageTwo.Top = $word.CentimetersToPoints(28.4)
$shapeBotPageTwo.Fill.Visible = 0
$shapeBotPageTwo.Line.Weight = 1
$shapeBotPageTwo.Line.Visible = 0
$shapeBotPageTwo.TextFrame.MarginBottom = $word.CentimetersToPoints(0)
$shapeBotPageTwo.TextFrame.MarginLeft = $word.CentimetersToPoints(0.1)
$shapeBotPageTwo.TextFrame.MarginRight = $word.CentimetersToPoints(0.1)
$shapeBotPageTwo.TextFrame.MarginTop = $word.CentimetersToPoints(0)

for ($i = 0; $i -lt 29; $i++) {
$table.Cell(15, 1).Delete()
}


#Отключить фиксированную высоту для ячейки, которая содержит перечень фдокументов и программ
$table.Cell(15, 1).HeightRule = 1
#Вставить таблицы со списками для аннлирования, замены и публикации
$document.Tables.Item(1).Cell(15, 1).TopPadding = $word.CentimetersToPoints(0.2)
$document.Tables.Item(1).Cell(15, 1).BottomPadding = $word.CentimetersToPoints(0.2)
$document.Tables.Item(1).Cell(15, 1).Range = [char]10 + [char]10 + "Заменить:" + [char]10 + [char]10 + [char]10 + [char]10 + "Аннулировать:" + [char]10 + [char]10 + [char]10 + [char]10 + "Выпустить:" + [char]10 + [char]10 + [char]10  + [char]10
$document.Tables.Item(1).Cell(15, 1).Range.Paragraphs.Item(4).Range.Select()
$document.Tables.Add($word.Selection.Range, 1, 1)
$document.Tables.Item(1).Cell(15, 1).Range.Paragraphs.Item(9).Range.Select()
$document.Tables.Add($word.Selection.Range, 1, 1)
$document.Tables.Item(1).Cell(15, 1).Range.Paragraphs.Item(14).Range.Select()
$document.Tables.Add($word.Selection.Range, 1, 1)
$document.Tables.Item(1).Cell(15, 1).Range.ParagraphFormat.Alignment = 1
Apply-FormattingInListTable -TableObject $document.Tables.Item(1).Tables.Item(1) -WordApp $word
Apply-FormattingInListTable -TableObject $document.Tables.Item(1).Tables.Item(2) -WordApp $word
Apply-FormattingInListTable -TableObject $document.Tables.Item(1).Tables.Item(3) -WordApp $word

Add-TestData -TableObject $document.Tables.Item(1).Tables.Item(1)
Add-TestData -TableObject $document.Tables.Item(1).Tables.Item(2)
Add-TestData -TableObject $document.Tables.Item(1).Tables.Item(3)
#Вставить поля
$table.Cell(5, 4).Range.Select()
$document.Application.Selection.Collapse(1)
$myField = $document.Fields.Add($document.Application.Selection.Range, 26)
$table.Cell(5, 4).Range.ParagraphFormat.Alignment = 1

$headerTable.Cell(2, 4).Range.Select()
$document.Application.Selection.Collapse(1)
$myField = $document.Fields.Add($document.Application.Selection.Range, 33)

$table.Cell(5, 3).Range.Select()
$document.Application.Selection.Collapse(1)
$myField = $document.Fields.Add($document.Application.Selection.Range, 33)
$table.Cell(5, 3).Range.ParagraphFormat.Alignment = 1
