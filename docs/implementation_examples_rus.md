# Примеры реализации для класса VBA Acordion

## Введение

Этот документ содержит различные примеры реализации класса `clsAccordion` в VBA. Примеры охватывают базовые и продвинутые сценарии использования, демонстрируя гибкость и функциональность класса.

## Пример 1: Простой аккордеон на форме

### Описание
Создание простого аккордеона с тремя элементами на пользовательской форме.

### Код
```vba
Sub CreateSimpleAccordion()
    Dim accordion As clsAccordion
    Set accordion = New clsAccordion
    accordion.SetParentForm Me
    
    ' Добавляем элементы аккордеона
    accordion.AddItem "Введение", "Это вводная информация о проекте"
    accordion.AddItem "Функции", "Список основных функций приложения"
    accordion.AddItem "Контакты", "Информация для связи с разработчиком"
    
    ' Создаем элементы управления на форме
    accordion.CreateControls
End Sub
```

## Пример 2: Аккордеон с настройкой стиля

### Описание
Создание аккордеона с настроенной цветовой схемой и высотой заголовков.

### Код
```vba
Sub CreateStyledAccordion()
    Dim accordion As clsAccordion
    Set accordion = New clsAccordion
    accordion.SetParentForm Me
    
    ' Настройка внешнего вида
    accordion.HeaderHeight = 35
    accordion.SetStyle RGB(65, 105, 225), RGB(248, 248, 255) ' Синие заголовки, почти белое содержимое
    accordion.SetAnimation 6 ' Средняя скорость анимации
    
    ' Добавляем элементы
    accordion.AddItem "Настройки интерфейса", "Параметры внешнего вида приложения"
    accordion.AddItem "Настройки безопасности", "Параметры защиты данных"
    accordion.AddItem "Настройки производительности", "Параметры оптимизации работы"
    
    accordion.CreateControls
End Sub
```

## Пример 3: Динамическое добавление элементов

### Описание
Пример добавления элементов аккордеона во время выполнения программы на основе данных из массива.

### Код
```vba
Sub CreateDynamicAccordion()
    Dim accordion As clsAccordion
    Set accordion = New clsAccordion
    accordion.SetParentForm Me
    
    ' Массив с данными для аккордеона
    Dim headers(1 To 3) As String
    Dim contents(1 To 3) As String
    
    headers(1) = "Пункт 1"
    contents(1) = "Содержимое первого пункта"
    headers(2) = "Пункт 2"
    contents(2) = "Содержимое второго пункта"
    headers(3) = "Пункт 3"
    contents(3) = "Содержимое третьего пункта"
    
    ' Добавляем элементы из массива
    Dim i As Integer
    For i = 1 To 3
        accordion.AddItem headers(i), contents(i)
    Next i
    
    accordion.CreateControls
End Sub
```

## Пример 4: Аккордеон с обработкой событий

### Описание
Создание аккордеона с обработкой событий открытия и закрытия элементов.

### Код
```vba
Sub CreateEventHandlingAccordion()
    Dim accordion As clsAccordion
    Set accordion = New clsAccordion
    accordion.SetParentForm Me
    
    ' Добавляем элементы
    accordion.AddItem "Данные пользователя", "Информация о текущем пользователе"
    accordion.AddItem "Настройки профиля", "Параметры настройки профиля"
    accordion.AddItem "История действий", "Журнал действий пользователя"
    
    ' Создаем элементы управления
    accordion.CreateControls
    
    ' Обработка событий (псевдокод - в реальной реализации потребуется дополнительная настройка)
    ' При раскрытии элемента
    ' Call accordion.OnItemExpanded(AddressOf HandleItemExpanded)
    ' При сворачивании элемента
    ' Call accordion.OnItemCollapsed(AddressOf HandleItemCollapsed)
End Sub

' Подпрограммы обработки событий
Sub HandleItemExpanded(itemIndex As Integer)
    Debug.Print "Элемент " & itemIndex & " раскрыт"
End Sub

Sub HandleItemCollapsed(itemIndex As Integer)
    Debug.Print "Элемент " & itemIndex & " свернут"
End Sub
```

## Пример 5: Аккордеон с вложенными элементами

### Описание
Создание многоуровневого аккордеона с возможностью вложения элементов друг в друга.

### Код
```vba
Sub CreateNestedAccordion()
    Dim mainaccordion As clsAccordion
    Set mainaccordion = New clsAccordion
    mainaccordion.SetParentForm Me
    
    ' Создаем основной аккордеон
    mainaccordion.AddItem "Категория 1", ""
    mainaccordion.AddItem "Категория 2", ""
    mainaccordion.AddItem "Категория 3", ""
    
    ' Для второй категории создаем вложенный аккордеон
    Dim nestedaccordion As clsAccordion
    Set nestedaccordion = New clsAccordion
    nestedaccordion.SetParentForm Me
    nestedaccordion.HeaderHeight = 25 ' Меньшая высота для вложенных элементов
    
    nestedaccordion.AddItem "Подкатегория 2.1", "Подробная информация о подкатегории 2.1"
    nestedaccordion.AddItem "Подкатегория 2.2", "Подробная информация о подкатегории 2.2"
    nestedaccordion.AddItem "Подкатегория 2.3", "Подробная информация о подкатегории 2.3"
    
    ' Вставляем вложенный аккордеон в содержимое второй категории
    ' Это требует дополнительной реализации в классе
    ' mainaccordion.Items(1).SetContentControl nestedaccordion
    
    mainaccordion.CreateControls
End Sub
```

## Пример 6: Аккордеон с кнопками действий

### Описание
Создание аккордеона, где каждый элемент содержит кнопки для выполнения действий.

### Код
```vba
Sub CreateActionAccordion()
    Dim accordion As clsAccordion
    Set accordion = New clsAccordion
    accordion.SetParentForm Me
    
    ' Добавляем элементы содержимым, включающим кнопки
    accordion.AddItem "Резервное копирование", "Создать резервную копию данных" & vbCrLf & "Кнопка: [Создать резерв]"
    accordion.AddItem "Очистка данных", "Очистить временные файлы" & vbCrLf & "Кнопка: [Очистить]"
    accordion.AddItem "Экспорт отчета", "Экспортировать отчет в Excel" & vbCrLf & "Кнопка: [Экспортировать]"
    
    accordion.CreateControls
    
    ' Добавляем обработчики для кнопок (требует дополнительной реализации)
    ' Это может потребовать модификации класса для поддержки встраивания кнопок
End Sub
```

## Пример 7: Аккордеон с сохранением состояния

### Описание
Создание аккордеона, который сохраняет и восстанавливает состояние (какие элементы были открыты/закрыты) между сеансами.

### Код
```vba
Sub CreateStatePreservingAccordion()
    Dim accordion As clsAccordion
    Set accordion = New clsAccordion
    accordion.SetParentForm Me
    
    ' Добавляем элементы
    accordion.AddItem "Настройки подключения", "Параметры подключения к базе данных"
    accordion.AddItem "Настройки отчетов", "Параметры формирования отчетов"
    accordion.AddItem "Настройки интерфейса", "Параметры внешнего вида приложения"
    
    ' Восстанавливаем состояние из сохраненных данных (псевдокод)
    ' Dim savedStates As Variant
    ' savedStates = GetSavedStates() ' Функция получения сохраненного состояния
    
    ' Применяем сохраненное состояние
    ' If IsArray(savedStates) Then
    '     Dim i As Integer
    '     For i = 0 To accordion.Items.Count - 1
    '         If i < UBound(savedStates) + 1 Then
    '             If savedStates(i) = True Then
    '                 accordion.Items(i).Expand
    '             Else
    '                 accordion.Items(i).Collapse
    '             End If
    '         End If
    '     Next i
    ' End If
    
    accordion.CreateControls
End Sub

' Подпрограмма сохранения состояния
Sub SaveAccordionState(accordion As clsAccordion)
    ' Сохраняем состояние каждого элемента
    Dim states() As Boolean
    ReDim states(0 To accordion.Items.Count - 1)
    
    Dim i As Integer
    For i = 0 To accordion.Items.Count - 1
        states(i) = accordion.Items(i).Expanded
    Next i
    
    ' Сохраняем массив в постоянное хранилище (например, в настройки приложения)
    ' SaveStatesToStorage states
End Sub
```

## Заключение

Эти примеры демонстрируют различные способы использования класса `clsAccordion` в VBA проектах. Вы можете адаптировать и комбинировать эти примеры в зависимости от ваших конкретных требований и сценариев использования.