' Класс аккордеона для управления элементами с возможностью открытия/закрытия
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 'True
END
Attribute VB_Name = "clsAcardion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Аккордеон для управления элементами с возможностью открытия/закрытия
Private mParentFrame As MSForms.Frame  ' Родительский фрейм, в котором размещаются элементы аккордеона
Private mItems      As Collection      ' Коллекция элементов аккордеона
Private mTopItem    As Single          ' Позиция по вертикали для следующего элемента
Private mItemCount  As Long            ' Количество элементов (используем Long вместо Byte для большего диапазона)
Private mCloseAfterOpen As Boolean     ' Флаг: закрывать другие элементы при открытии нового
Public mbScrolShow  As Boolean         ' Флаг: показывать ли полосу прокрутки

' Открывает или закрывает все элементы аккордеона
' Параметры:
'   bOpen - True для открытия, False для закрытия
Public Sub OpenAll(ByVal bOpen As Boolean)
    If mCloseAfterOpen Then Exit Sub ' Если включено автоматическое закрытие, не открываем все
    
    Dim i As Long ' Изменяем тип на Long для совместимости с возможным большим количеством элементов
    For i = 1 To Me.Count
        With Me.item(i)
            If .IsOpen <> bOpen And .Enabled = True Then
                .IsOpen = bOpen
            End If
        End With
    Next i
End Sub

' Инициализирует аккордеон с указанным родительским фреймом
' Параметры:
'   MainFame - родительский фрейм
'   CloseAfterOpen - флаг автоматического закрытия других элементов (по умолчанию True)
Public Sub Initialize(ByRef MainFame As MSForms.Frame, Optional CloseAfterOpen As Boolean = True)
    Set mParentFrame = MainFame
    Set mItems = New Collection
    mCloseAfterOpen = CloseAfterOpen
End Sub

' Возвращает количество элементов в аккордеоне
Public Property Get Count() As Long ' Изменяем тип на Long для большего диапазона
    Count = mItems.Count
End Property

' Возвращает коллекцию элементов аккордеона
Public Property Get Items() As Collection
    Set Items = mItems
End Property

' Возвращает элемент аккордеона по индексу
' Параметры:
'   index - индекс элемента (1-based)
Public Property Get item(ByVal index As Long) As clsAcardionItem ' Изменяем тип индекса на Long
    If index < 1 Or index > mItems.Count Then
        Set item = Nothing
        Exit Property
    End If
    Set item = mItems(index)
End Property

' Добавляет новый элемент в аккордеон
' Параметры:
'   sTitle - заголовок элемента
'   sBody - содержимое элемента
'   Locked - заблокировано ли содержимое для редактирования (по умолчанию False)
'   TitleHeight - высота заголовка (по умолчанию 20)
'   BodyHeight - высота тела элемента (по умолчанию 50)
'   ForeColorBtn - цвет текста кнопки (по умолчанию черный)
'   OpenCaption - символ для открытого состояния (по умолчанию 6)
'   CloseCaption - символ для закрытого состояния (по умолчанию 5)
Public Function AddItem(ByRef sTitle As String, ByRef sBody As String, Optional Locked As Boolean = False, _
        Optional TitleHeight As Single = 20, Optional BodyHeight As Single = 50, Optional ForeColorBtn As Long = vbBlack, _
        Optional OpenCaption As String = 6, Optional CloseCaption As String = 5) As clsAcardionItem
    Dim acordItem As clsAcardionItem
    Set acordItem = New clsAcardionItem
    mItemCount = mItemCount + 1
    Call acordItem.AddItem_(Me, mParentFrame, sTitle, sBody, mTopItem, mItems, mItemCount, _
            TitleHeight, BodyHeight, Locked, ForeColorBtn, mCloseAfterOpen, OpenCaption, CloseCaption)
    Call mItems.Add(acordItem, acordItem.Name)
    mTopItem = mTopItem + TitleHeight
    Set AddItem = acordItem
End Function

' Удаляет все элементы из аккордеона
Public Sub RemoveAll()
    ' Удаляем все элементы из родительского фрейма
    On Error Resume Next ' Обрабатываем возможную ошибку при очистке
    mParentFrame.Controls.Clear
    On Error GoTo 0
    
    ' Очищаем коллекцию элементов
    Set mItems = Nothing
    Set mItems = New Collection
    mTopItem = 0
    mItemCount = 0
End Sub

' Удаляет элемент из аккордеона по индексу
' Параметры:
'   index - индекс элемента для удаления (1-based)
' Возвращает True при успешном удалении, False в противном случае
Public Function RemoveItem(ByVal index As Long) As Boolean ' Изменяем тип индекса на Long
    Dim acordItem As clsAcardionItem
    Set acordItem = Me.item(index)
    If acordItem Is Nothing Then Exit Function ' Проверяем, существует ли элемент
    
    ' Удаляем элемент из коллекции и с формы
    With acordItem
        On Error Resume Next ' Обрабатываем возможную ошибку при удалении элемента управления
        Call .Parent.Controls.Remove(.ItemFrame.Name)
        On Error GoTo 0
        Call mItems.Remove(index)
        Call .moveControls_
    End With
    Set acordItem = Nothing
    RemoveItem = True
End Function

' Освобождает ресурсы при уничтожении объекта
Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mParentFrame = Nothing
End Sub
