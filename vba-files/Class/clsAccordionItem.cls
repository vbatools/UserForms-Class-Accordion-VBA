VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccordionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private variables for managing accordion item controls and state
Private mParentFrame As MSForms.Frame   ' Parent frame where the item is placed
Private mItemFrame  As MSForms.Frame    ' Frame for this item
Private mBodyTxtBox As MSForms.TextBox  ' Text box for item body
Private mItems      As Collection       ' Collection of all accordion items

Private mParent     As clsAccordion      ' Reference to parent accordion
Private mTitleHeight As Single          ' Height of title section
Private mBodyHeight As Single           ' Height of body section
Private mNameItem   As String           ' Name of item
Private mIsOpen     As Boolean          ' State: open/closed
Private mbCloseAfterOpen As Boolean     ' Flag: close other items when opening this one
Private mOpenCaption As String          ' Caption for open state
Private mCloseCaption As String         ' Caption for close state

Private WithEvents mTitleLabel As MSForms.Label  ' Title label
Attribute mTitleLabel.VB_VarHelpID = -1
Private WithEvents mTitleBtn As MSForms.Label    ' Title button
Attribute mTitleBtn.VB_VarHelpID = -1

' Get close caption
Public Property Get CloseCaption() As String
    CloseCaption = mCloseCaption
End Property

' Set close caption
Public Property Let CloseCaption(ByVal Caption As String)
    mCloseCaption = Caption
End Property

' Get open caption
Public Property Get OpenCaption() As String
    OpenCaption = mOpenCaption
End Property

' Set open caption
Public Property Let OpenCaption(ByVal Caption As String)
    mOpenCaption = Caption
End Property

' Get item name
Public Property Get Name() As String
    Name = mNameItem
End Property

' Get item frame
Public Property Get ItemFrame() As MSForms.Frame
    Set ItemFrame = mItemFrame
End Property

' Get parent frame
Public Property Get Parent() As MSForms.Frame
    Set Parent = mParentFrame
End Property

' Get body text box
Public Property Get BodyTxtBox() As MSForms.TextBox
    Set BodyTxtBox = mBodyTxtBox
End Property

' Get title label
Public Property Get TitleLabel() As MSForms.Label
    Set TitleLabel = mTitleLabel
End Property

' Get title button
Public Property Get TitleBtn() As MSForms.Label
    Set TitleBtn = mTitleBtn
End Property

' Get open state (open/closed)
Public Property Get IsOpen() As Boolean
    IsOpen = mIsOpen
End Property

' Set open state (open/closed)
Public Property Let IsOpen(ByVal bOpen As Boolean)
    Call setClick
End Property

' Get enabled state
Public Property Get Enabled() As Boolean
    Enabled = mTitleBtn.Enabled
End Property

' Set enabled state
Public Property Let Enabled(ByVal bEnabled As Boolean)
    mTitleBtn.Enabled = bEnabled
End Property

' Set background color
Public Property Let BackColor(ByVal Color As Long)
    mTitleBtn.BackColor = Color
    mTitleLabel.BackColor = Color
End Property

' Get foreground color (returns title label color if both have same color)
Public Property Get ForeColor() As Long
    If mTitleBtn.ForeColor = mTitleLabel.ForeColor Then ForeColor = mTitleLabel.ForeColor
End Property

' Set foreground color
Public Property Let ForeColor(ByVal Color As Long)
    mTitleBtn.ForeColor = Color
    mTitleLabel.ForeColor = Color
End Property

' Get background color (returns title label color if both have same color)
Public Property Get BackColor() As Long
    If mTitleBtn.BackColor = mTitleLabel.BackColor Then BackColor = mTitleLabel.BackColor
End Property

' Move controls to update positions when items open/close
Public Sub moveControls_()
    Dim item As clsAccordionItem
    Dim sngTop As Single
    sngTop = 0
    
    ' Loop through all items
    For Each item In mItems
        With item
            ' If auto-close is enabled, close other open items
            If mbCloseAfterOpen Then
                If .IsOpen And mNameItem <> .Name Then
                    ' Close item (update height and caption)
                    .ItemFrame.height = mTitleHeight
                    .TitleBtn.Caption = mCloseCaption
                    .IsOpen = False ' Update internal state
                End If
            End If
            ' Update position of item frame
            With .ItemFrame
                .Top = sngTop
                sngTop = .Top + .height
            End With
        End With
    Next item
End Sub

' Initialize accordion item
' Parameters:
'   Parent - Parent accordion
'   ParentFrame - Parent frame
'   sTitle - Title text
'   sBody - Body content
'   Top - Top position for item
'   coll - Collection of items
'   Name - Name of item
'   TitleHeight - Height of title section
'   BodyHeight - Height of body section
'   LockedBoby - Whether body is locked for editing
'   ForeColorBtn - Color of button text
'   CloseAfterOpen - Whether to close other items when opening this one
'   OpenCaption - Character for open state
'   CloseCaption - Character for close state
Public Sub AddItem_(ByRef Parent As clsAccordion, ByRef ParentFrame As MSForms.Frame, ByRef sTitle As String, ByRef sBody As String, ByRef Top As Single, _
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
        .Caption = sTitle
    End With
    With mTitleBtn
        .left = mTitleLabel.left + mTitleLabel.width
        .width = mItemFrame.width * 0.1
        .height = mItemFrame.height
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

' Handle click on title label
Private Sub mTitleLabel_Click()
    Call setClick
End Sub

' Handle click on title button
Private Sub mTitleBtn_Click()
    Call setClick
End Sub

' Toggle open state and update UI
Private Sub setClick()
    ' Update state
    If mIsOpen Then
        mTitleBtn.Caption = mCloseCaption
        mItemFrame.height = mTitleHeight
    Else
        mTitleBtn.Caption = mOpenCaption
        mItemFrame.height = mTitleHeight + mBodyHeight
    End If
    
    ' Update positions of other controls
    Call moveControls_
    mIsOpen = Not mIsOpen
    
    ' Update scrollbar settings
    With mParentFrame
        Dim snHeight As Single
        ' Calculate total height needed for all items
        snHeight = mItems(mItems.Count).ItemFrame.Top + mItems(mItems.Count).ItemFrame.height
        If .height <= snHeight Then
            ' If content exceeds frame height, show scrollbar
            .ScrollHeight = snHeight
            .ScrollBars = fmScrollBarsVertical
            Call .Scroll(0, fmScrollActionLineUp)
        Else
            ' Otherwise hide scrollbar
            Call .Scroll(0, fmScrollActionBegin)
            .ScrollHeight = .height
            .ScrollBars = fmScrollBarsNone
        End If
        ' Update width of controls based on scrollbar visibility
        Call changeWidth(.width, .ScrollBars)
    End With
End Sub

' Update width of controls based on frame width and scrollbar visibility
' Parameters:
'   snWidth - Frame width
'   ScrolShow - Scrollbar visibility
Private Sub changeWidth(ByRef snWidth As Single, ByRef ScrolShow As fmScrollBars)
    Dim delta As Integer
    delta = 15 ' Width of scrollbar
    
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

' Cleanup when object is terminated
Private Sub Class_Terminate()
    Set mItemFrame = Nothing
    Set mTitleLabel = Nothing
    Set mBodyTxtBox = Nothing
    Set mTitleBtn = Nothing
End Sub

