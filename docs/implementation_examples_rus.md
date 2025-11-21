# Примеры реализации для класса VBA Acordion

## Введение

Этот документ содержит различные примеры реализации класса `clsAcardion` в VBA. Примеры охватывают базовые и продвинутые сценарии использования, демонстрируя гибкость и функциональность класса.

## Пример 1: Простой аккордеон на форме

### Описание
Создание простого аккордеона с тремя элементами на пользовательской форме.

### Код
```vba
Sub CreateSimpleAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Добавляем элементы аккордеона
    acardion.AddItem "Введение", "Это вводная информация о проекте"
    acardion.AddItem "Функции", "Список основных функций приложения"
    acardion.AddItem "Контакты", "Информация для связи с разработчиком"
    
    ' Создаем элементы управления на форме
    acardion.CreateControls
End Sub
```

## Пример 2: Аккордеон с настройкой стиля

### Описание
Создание аккордеона с настроенной цветовой схемой и высотой заголовков.

### Код
```vba
Sub CreateStyledAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Настройка внешнего вида
    acardion.HeaderHeight = 35
    acardion.SetStyle RGB(65, 105, 225), RGB(248, 248, 255) ' Синие заголовки, почти белое содержимое
    acardion.SetAnimation 6 ' Средняя скорость анимации
    
    ' Добавляем элементы
    acardion.AddItem "Настройки интерфейса", "Параметры внешнего вида приложения"
    acardion.AddItem "Настройки безопасности", "Параметры защиты данных"
    acardion.AddItem "Настройки производительности", "Параметры оптимизации работы"
    
    acardion.CreateControls
End Sub
```

## Пример 3: Динамическое добавление элементов

### Описание
Пример добавления элементов аккордеона во время выполнения программы на основе данных из массива.

### Код
```vba
Sub CreateDynamicAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
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
        acardion.AddItem headers(i), contents(i)
    Next i
    
    acardion.CreateControls
End Sub
```

## Пример 4: Аккордеон с обработкой событий

### Описание
Создание аккордеона с обработкой событий открытия и закрытия элементов.

### Код
```vba
Sub CreateEventHandlingAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Добавляем элементы
    acardion.AddItem "Данные пользователя", "Информация о текущем пользователе"
    acardion.AddItem "Настройки профиля", "Параметры настройки профиля"
    acardion.AddItem "История действий", "Журнал действий пользователя"
    
    ' Создаем элементы управления
    acardion.CreateControls
    
    ' Обработка событий (псевдокод - в реальной реализации потребуется дополнительная настройка)
    ' При раскрытии элемента
    ' Call acardion.OnItemExpanded(AddressOf HandleItemExpanded)
    ' При сворачивании элемента
    ' Call acardion.OnItemCollapsed(AddressOf HandleItemCollapsed)
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
    Dim mainAcardion As New clsAcardion
    mainAcardion.SetParentForm Me
    
    ' Создаем основной аккордеон
    mainAcardion.AddItem "Категория 1", ""
    mainAcardion.AddItem "Категория 2", ""
    mainAcardion.AddItem "Категория 3", ""
    
    ' Для второй категории создаем вложенный аккордеон
    Dim nestedAcardion As New clsAcardion
    nestedAcardion.SetParentForm Me
    nestedAcardion.HeaderHeight = 25 ' Меньшая высота для вложенных элементов
    
    nestedAcardion.AddItem "Подкатегория 2.1", "Подробная информация о подкатегории 2.1"
    nestedAcardion.AddItem "Подкатегория 2.2", "Подробная информация о подкатегории 2.2"
    nestedAcardion.AddItem "Подкатегория 2.3", "Подробная информация о подкатегории 2.3"
    
    ' Вставляем вложенный аккордеон в содержимое второй категории
    ' Это требует дополнительной реализации в классе
    ' mainAcardion.Items(1).SetContentControl nestedAcardion
    
    mainAcardion.CreateControls
End Sub
```

## Пример 6: Аккордеон с кнопками действий

### Описание
Создание аккордеона, где каждый элемент содержит кнопки для выполнения действий.

### Код
```vba
Sub CreateActionAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Добавляем элементы с содержимым, включающим кнопки
    acardion.AddItem "Резервное копирование", "Создать резервную копию данных" & vbCrLf & "Кнопка: [Создать резерв]"
    acardion.AddItem "Очистка данных", "Очистить временные файлы" & vbCrLf & "Кнопка: [Очистить]"
    acardion.AddItem "Экспорт отчета", "Экспортировать отчет в Excel" & vbCrLf & "Кнопка: [Экспортировать]"
    
    acardion.CreateControls
    
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
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Добавляем элементы
    acardion.AddItem "Настройки подключения", "Параметры подключения к базе данных"
    acardion.AddItem "Настройки отчетов", "Параметры формирования отчетов"
    acardion.AddItem "Настройки интерфейса", "Параметры внешнего вида приложения"
    
    ' Восстанавливаем состояние из сохраненных данных (псевдокод)
    ' Dim savedStates As Variant
    ' savedStates = GetSavedStates() ' Функция получения сохраненного состояния
    
    ' Применяем сохраненное состояние
    ' If IsArray(savedStates) Then
    '     Dim i As Integer
    '     For i = 0 To acardion.Items.Count - 1
    '         If i < UBound(savedStates) + 1 Then
    '             If savedStates(i) = True Then
    '                 acardion.Items(i).Expand
    '             Else
    '                 acardion.Items(i).Collapse
    '             End If
    '         End If
    '     Next i
    ' End If
    
    acardion.CreateControls
End Sub

' Подпрограмма сохранения состояния
Sub SaveAccordionState(acardion As clsAcardion)
    ' Сохраняем состояние каждого элемента
    Dim states() As Boolean
    ReDim states(0 To acardion.Items.Count - 1)
    
    Dim i As Integer
    For i = 0 To acardion.Items.Count - 1
        states(i) = acardion.Items(i).Expanded
    Next i
    
    ' Сохраняем массив в постоянное хранилище (например, в настройки приложения)
    ' SaveStatesToStorage states
End Sub
```

## Заключение

Эти примеры демонстрируют различные способы использования класса `clsAcardion` в VBA проектах. Вы можете адаптировать и комбинировать эти примеры в зависимости от ваших конкретных требований и сценариев использования.