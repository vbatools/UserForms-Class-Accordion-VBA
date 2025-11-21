# Техническая документация для класса VBA Accordion

## Обзор

Класс `clsAccordion` реализует функциональность аккордеона в VBA. Он позволяет создавать интерактивные элементы управления, которые могут раскрываться и сворачиваться. Класс состоит из основного класса `clsAccordion` и вспомогательного класса `clsAccordionItem`, представляющего отдельный элемент аккордеона.

## Структура класса

### clsAccordion
- `ParentForm` - ссылка на родительскую форму
- `Items` - коллекция элементов аккордеона
- `HeaderHeight` - высота заголовков
- `AnimationSpeed` - скорость анимации
- `HeaderColor` - цвет заголовков
- `ContentColor` - цвет содержимого

### clsAccordionItem
- `Header` - текст заголовка
- `Content` - содержимое элемента
- `Expanded` - состояние (развернут/свернут)
- `HeaderControl` - элемент управления заголовка
- `ContentControl` - элемент управления содержимого

## Методы

### clsAccordion
- `AddItem(Header As String, Content As String)` - добавляет новый элемент
- `RemoveItem(Index As Integer)` - удаляет элемент по индексу
- `ClearItems()` - очищает все элементы
- `CreateControls()` - создает элементы управления на форме
- `ExpandAll()` - разворачивает все элементы
- `CollapseAll()` - сворачивает все элементы
- `SetStyle(HeaderColor As Long, ContentColor As Long)` - устанавливает стили
- `SetAnimation(Speed As Integer)` - устанавливает скорость анимации

### clsAccordionItem
- `SetHeader(Header As String)` - устанавливает текст заголовка
- `SetContent(Content As String)` - устанавливает содержимое
- `Expand()` - разворачивает элемент
- `Collapse()` - сворачивает элемент
- `Toggle()` - переключает состояние элемента
- `UpdateLayout()` - обновляет размещение элементов

## Использование

1. Создайте экземпляр класса `clsAccordion`
2. Установите родительскую форму с помощью `SetParentForm`
3. Добавьте элементы с помощью `AddItem`
4. Вызовите `CreateControls` для создания элементов управления на форме
5. Используйте методы `ExpandAll`, `CollapseAll` для управления состоянием

## Пример кода

```vba
Dim accordion As clsAccordion
Set accordion = New clsAccordion
accordion.SetParentForm Me
accordion.AddItem "Элемент 1", "Содержимое первого элемента"
accordion.AddItem "Элемент 2", "Содержимое второго элемента"
accordion.CreateControls
```

## События

Класс поддерживает события:
- `OnItemExpanded` - при раскрытии элемента
- `OnItemCollapsed` - при сворачивании элемента
- `OnAllExpanded` - когда все элементы раскрыты
- `OnAllCollapsed` - когда все элементы свернуты