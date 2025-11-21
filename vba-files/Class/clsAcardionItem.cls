' Класс элемента аккордеона с заголовком и телом
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 'True
END
Attribute VB_Name = "clsAcardionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Элемент аккордеона с заголовком и телом
Private mParentFrame As MSForms.Frame   ' Родительский фрейм, в котором размещается элемент
Private mItemFrame  As MSForms.Frame    ' Фрейм элемента
Private mBodyTxtBox As MSForms.TextBox  ' Текстовое поле тела элемента
Private mItems      As Collection       ' Коллекция всех элементов аккордеона

Private mParent     As clsAcardion      ' Ссылка на родительский аккордеон
Private mTitleHeight As Single          ' Высота заголовка
Private mBodyHeight As Single           ' Высота тела элемента
Private mNameItem   As String           ' Имя элемента
Private mIsOpen     As Boolean          ' Состояние: открыт/закрыт
Private mbCloseAfterOpen As Boolean     ' Флаг: закрывать другие элементы при открытии этого
Private mOpenCaption As String          ' Заголовок для открытого состояния
Private mCloseCaption As String         ' Заголовок для закрытого состояния
'Private mVisible    As Boolean

Private WithEvents mTitleLabel As MSForms.Label  ' Метка заголовка
Attribute mTitleLabel.VB_VarHelpID = -1
Private WithEvents mTitleBtn As MSForms.Label    ' Кнопка заголовка
Attribute mTitleBtn.VB_VarHelpID = -1

' Возвращает заголовок для закрытого состояния
Public Property Get CloseCaption() As String
    CloseCaption = mCloseCaption
End Property

' Устанавливает заголовок для закрытого состояния
Public Property Let CloseCaption(ByVal Caption As String)
    mCloseCaption = Caption
End Property

' Возвращает заголовок для открытого состояния
Public Property Get OpenCaption() As String
    OpenCaption = mOpenCaption
End Property

' Устанавливает заголовок для открытого состояния
Public Property Let OpenCaption(ByVal Caption As String)
    mOpenCaption = Caption
End Property

' Возвращает имя элемента
Public Property Get Name() As String
    Name = mNameItem
End Property

' Возвращает фрейм элемента
Public Property Get ItemFrame() As MSForms.Frame
    Set ItemFrame = mItemFrame
End Property

' Возвращает родительский фрейм
Public Property Get Parent() As MSForms.Frame
    Set Parent = mParentFrame
End Property

' Возвращает текстовое поле тела элемента
Public Property Get BodyTxtBox() As MSForms.TextBox
    Set BodyTxtBox = mBodyTxtBox
End Property

' Возвращает метку заголовка
Public Property Get TitleLabel() As MSForms.Label
    Set TitleLabel = mTitleLabel
End Property

' Возвращает кнопку заголовка
Public Property Get TitleBtn() As MSForms.Label
    Set TitleBtn = mTitleBtn
End Property

' Возвращает состояние элемента (открыт/закрыт)
Public Property Get IsOpen() As Boolean
    IsOpen = mIsOpen
End Property

' Устанавливает состояние элемента (открыт/закрыт)
Public Property Let IsOpen(ByVal bOpen As Boolean)
    Call setClick
End Property

' Возвращает состояние доступности элемента
Public Property Get Enabled() As Boolean
    Enabled = mTitleBtn.Enabled
End Property

' Устанавливает состояние доступности элемента
Public Property Let Enabled(ByVal bEnabled As Boolean)
    mTitleBtn.Enabled = bEnabled
End Property

' Устанавливает цвет фона
Public Property Let BackColor(ByVal Color As Long)
    mTitleBtn.BackColor = Color
    mTitleLabel.BackColor = Color
End Property

' Возвращает цвет текста (только для чтения, так как управляется через ForeColor=)
Public Property Get ForeColor() As Long
    If mTitleBtn.ForeColor = mTitleLabel.ForeColor Then ForeColor = mTitleLabel.ForeColor
End Property

' Устанавливает цвет текста
Public Property Let ForeColor(ByVal Color As Long)
    mTitleBtn.ForeColor = Color
    mTitleLabel.ForeColor = Color
End Property

' Возвращает цвет фона (только для чтения, так как управляется через BackColor=)
Public Property Get BackColor() As Long
    If mTitleBtn.BackColor = mTitleLabel.BackColor Then BackColor = mTitleLabel.BackColor
End Property

'Public Property Get Visible() As Boolean
'    Visible = mVisible
'End Property
'
'Public Property Let Visible(ByVal bVisible As Boolean)
'    mVisible = bVisible
'    If mVisible Then
'        If mIsOpen Then
'            mTitleBtn.Caption = mCloseCaption
'            mItemFrame.Height = mTitleHeight
'        Else
'            mTitleBtn.Caption = mOpenCaption
'            mItemFrame.Height = mTitleHeight + mBodyHeight
'        End If
'    Else
'        mItemFrame.Height = 0
'    End If
'End Property

' Обновляет положение и состояние всех элементов аккордеона
Public Sub moveControls_()
    Dim item As clsAcardionItem
    Dim sngTop As Single
    sngTop = 0
    
    ' Обновляем положение всех элементов
    For Each item In mItems
        With item
            ' Если включено автоматическое закрытие, закрываем другие элементы
            If mbCloseAfterOpen Then
                If .IsOpen And mNameItem <> .Name Then
                    ' Закрываем элемент (устанавливаем высоту заголовка и обновляем заголовок кнопки)
                    .ItemFrame.height = mTitleHeight
                    .TitleBtn.Caption = mCloseCaption
                    .IsOpen = False ' Устанавливаем состояние как закрытое
                End If
            End If
            ' Устанавливаем вертикальное положение элемента
            With .ItemFrame
                .Top = sngTop
                sngTop = .Top + .height
            End With
        End With
    Next item
End Sub

' Инициализирует элемент аккордеона
' Параметры:
'   Parent - родительский аккордеон
'   ParentFrame - родительский фрейм
'   sTitle - заголовок элемента
'   sBody - содержимое элемента
'   Top - начальная позиция по вертикали
'   coll - коллекция элементов
'   Name - имя элемента
'   TitleHeight - высота заголовка
'   BodyHeight - высота тела элемента
'   LockedBoby - заблокировано ли тело для редактирования
'   ForeColorBtn - цвет текста кнопки
'   CloseAfterOpen - флаг автоматического закрытия других элементов
'   OpenCaption - символ для открытого состояния
'   CloseCaption - символ для закрытого состояния
Public Sub AddItem_(ByRef Parent As clsAcardion, ByRef ParentFrame As MSForms.Frame, ByRef sTitle As String, ByRef sBody As String, ByRef Top As Single, _
        ByRef coll As Collection, ByVal Name As String, ByVal TitleHeight As Single, ByVal BodyHeight As Single, _
        ByRef LockedBoby As Boolean, ByRef ForeColorBtn As Long, ByRef CloseAfterOpen As Boolean, _
        ByRef OpenCaption As String, ByRef CloseCaption As String)

    Set mParentFrame = ParentFrame
    mTitleHeight = TitleHeight
    mBodyHeight = BodyHeight
    mOpenCaption = OpenCaption
    mCloseCaption = CloseCaption
    With mParentFrame.Controls
        mNameItem = Name
        Set mItemFrame = .Add("Forms.Frame.1", mNameItem, True)
    End With
    With mItemFrame
        .Top = Top
        .width = mParentFrame.width
        .height = mTitleHeight
        Set mTitleLabel = .Add("Forms.Label.1", "title_" & mNameItem, True)
        Set mTitleBtn = .Add("Forms.Label.1", "btn_" & mNameItem, True)
        Set mBodyTxtBox = .Add("Forms.TextBox.1", "body_" & mNameItem, True)
    End With
    With mTitleLabel
        .width = mItemFrame.width * 0.9
        .height = mTitleHeight
        '.BorderStyle = fmBorderStyleSingle
        .Caption = sTitle
    End With
    With mTitleBtn
        .left = mTitleLabel.left + mTitleLabel.width
        .width = mItemFrame.width * 0.1
        .height = mItemFrame.height
        '.BorderStyle = fmBorderStyleSingle
        .Font.Size = 15
        .Font.Bold = True
        .Font.Name = "Webdings"
        .TextAlign = fmTextAlignCenter
        .Caption = mCloseCaption
        .ForeColor = ForeColorBtn
    End With
    With mBodyTxtBox
        .width = mItemFrame.width
        .Top = mTitleLabel.Top + mTitleLabel.height
        .height = mBodyHeight
        .BorderStyle = fmBorderStyleSingle
        .MultiLine = True
        .Value = sBody
        .Locked = LockedBoby
    End With
    mIsOpen = False
    Set mItems = coll
    mbCloseAfterOpen = CloseAfterOpen
    Set mParent = Parent
End Sub

' Обработчик клика по метке заголовка
Private Sub mTitleLabel_Click()
    Call setClick
End Sub

' Обработчик клика по кнопке заголовка
Private Sub mTitleBtn_Click()
    Call setClick
End Sub

' Переключает состояние элемента (открыт/закрыт) и обновляет UI
Private Sub setClick()
    ' Переключаем состояние
    If mIsOpen Then
        mTitleBtn.Caption = mCloseCaption
        mItemFrame.height = mTitleHeight
    Else
        mTitleBtn.Caption = mOpenCaption
        mItemFrame.height = mTitleHeight + mBodyHeight
    End If
    
    ' Обновляем расположение элементов
    Call moveControls_
    mIsOpen = Not mIsOpen
    
    ' Обновляем прокрутку родительского фрейма
    With mParentFrame
        Dim snHeight As Single
        ' Вычисляем общую высоту всех элементов
        snHeight = mItems(mItems.Count).ItemFrame.Top + mItems(mItems.Count).ItemFrame.height
        If .height <= snHeight Then
            ' Если содержимое больше фрейма, включаем прокрутку
            .ScrollHeight = snHeight
            .ScrollBars = fmScrollBarsVertical
            Call .Scroll(0, fmScrollActionLineUp)
        Else
            ' Иначе отключаем прокрутку
            Call .Scroll(0, fmScrollActionBegin)
            .ScrollHeight = .height
            .ScrollBars = fmScrollBarsNone
        End If
        ' Обновляем ширину элементов в зависимости от наличия полосы прокрутки
        Call changeWidth(.width, .ScrollBars)
    End With
End Sub

' Обновляет ширину всех элементов с учетом полосы прокрутки
' Параметры:
'   snWidth - новая ширина
'   ScrolShow - тип полосы прокрутки
Private Sub changeWidth(ByRef snWidth As Single, ByRef ScrolShow As fmScrollBars)
    Dim delta As Integer
    delta = 15 ' Ширина полосы прокрутки
    
    If ScrolShow = fmScrollBarsNone Then
        delta = 0
        mParent.mbScrolShow = False
    Else
        If mParent.mbScrolShow Then Exit Sub
        mParent.mbScrolShow = True
    End If
    
    Dim i As Long
    For i = 1 To mItems.Count
        With mItems(i)
            .ItemFrame.width = snWidth - delta
            .BodyTxtBox.width = snWidth - delta
            .TitleBtn.left = .ItemFrame.width * 0.9
        End With
    Next i
End Sub

' Освобождает ресурсы при уничтожении объекта
Private Sub Class_Terminate()
    Set mItemFrame = Nothing
    Set mTitleLabel = Nothing
    Set mBodyTxtBox = Nothing
    Set mTitleBtn = Nothing
End Sub
