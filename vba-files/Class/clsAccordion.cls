VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccordion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private variables for managing accordion control elements and state
Private mParentFrame As MSForms.Frame  ' Parent frame where accordion items are placed
Private mItems      As Collection      ' Collection of accordion items
Private mTopItem    As Single          ' Top position for next item
Private mItemCount  As Long            ' Count of items (using Long instead of Byte for larger capacity)
Private mCloseAfterOpen As Boolean     ' Flag: close other items when opening a new one
Public mbScrolShow  As Boolean         ' Flag: scrollbar is visible

' Open or close all items
' Parameters:
'   bOpen - True to open, False to close
Public Sub OpenAll(ByVal bOpen As Boolean)
    If mCloseAfterOpen Then Exit Sub    ' If auto-close is enabled, skip this operation

    Dim i           As Long    ' Loop variable using Long for compatibility with collection indices
    For i = 1 To Me.Count
        With Me.item(i)
            ' Only process visible items
            If .Visible And .Enabled = True Then
                If .IsOpen <> bOpen Then
                    .IsOpen = bOpen
                End If
            End If
        End With
    Next i
End Sub

' Initialize accordion with parent frame
' Parameters:
'   MainFame - Parent frame
'   CloseAfterOpen - Whether to close other items when opening a new one (default True)
Public Sub Initialize(ByRef MainFame As MSForms.Frame, Optional CloseAfterOpen As Boolean = True)
    Set mParentFrame = MainFame
    Set mItems = New Collection
    mCloseAfterOpen = CloseAfterOpen
End Sub

' Get count of items
Public Property Get Count() As Long    ' Using Long for count
    Count = mItems.Count
End Property

' Get collection of items
Public Property Get Items() As Collection
    Set Items = mItems
End Property

' Get item by index
' Parameters:
'   index - Item index (1-based)
Public Property Get item(ByVal index As Long) As clsAccordionItem    ' Using Long for index
    If index < 1 Or index > mItems.Count Then
        Set item = Nothing
        Exit Property
    End If
    Set item = mItems(index)
End Property

' Add new accordion item
' Parameters:
'   sTitle - Title text
'   sBody - Body content
'   Locked - Whether body is locked for editing (default False)
'   TitleHeight - Height of title section (default 20)
'   BodyHeight - Height of body section (default 50)
'   ForeColorBtn - Color of button text (default black)
'   OpenCaption - Character for open state (default 6)
'   CloseCaption - Character for close state (default 5)
Public Function AddItem(ByRef sTitle As String, ByRef sBody As String, Optional Locked As Boolean = False, _
        Optional TitleHeight As Single = 20, Optional BodyHeight As Single = 50, Optional ForeColorBtn As Long = vbBlack, _
        Optional OpenCaption As String = 6, Optional CloseCaption As String = 5) As clsAccordionItem
    Dim acordItem   As clsAccordionItem
    Set acordItem = New clsAccordionItem
    mItemCount = mItemCount + 1
    Call acordItem.AddItem_(Me, mParentFrame, sTitle, sBody, mTopItem, mItems, mItemCount, _
            TitleHeight, BodyHeight, Locked, ForeColorBtn, mCloseAfterOpen, OpenCaption, CloseCaption)
    Call mItems.Add(acordItem, acordItem.Name)
    mTopItem = mTopItem + TitleHeight
    ' Update positions of all items after adding new one
    Call moveControls_
    Set AddItem = acordItem
End Function

' Remove all items from accordion
Public Sub RemoveAll()
    ' Clear all controls from parent frame
    On Error Resume Next    ' Suppress errors during cleanup
    mParentFrame.Controls.Clear
    On Error GoTo 0

    ' Reset collection and counters
    Set mItems = Nothing
    Set mItems = New Collection
    mTopItem = 0
    mItemCount = 0
    ' Update positions after clearing all items
    Call moveControls_
End Sub

' Remove item by index
' Parameters:
'   index - Index of item to remove (1-based)
' Returns True if successful, False otherwise
Public Function RemoveItem(ByVal index As Long) As Boolean    ' Using Long for index
    Dim acordItem   As clsAccordionItem
    Set acordItem = Me.item(index)
    If acordItem Is Nothing Then Exit Function    ' Exit if item doesn't exist

    ' Remove item from collection and controls
    With acordItem
        On Error Resume Next    ' Suppress errors when removing controls
        Call .Parent.Controls.Remove(.ItemFrame.Name)
        On Error GoTo 0
        Call mItems.Remove(index)
        ' Update positions of remaining items after removal
        Call moveControls_
    End With
    Set acordItem = Nothing
    RemoveItem = True
End Function

' Move controls to update positions when items are added/removed
Private Sub moveControls_()
    ' Call moveControls_ for the first item to update all positions
    If mItems.Count > 0 Then
        Dim firstItem As clsAccordionItem
        Set firstItem = mItems(1)
        Call firstItem.moveControls_
    End If
End Sub

' Cleanup when object is terminated
Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mParentFrame = Nothing
End Sub
