Attribute VB_Name = "modShowForms"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Module1
'* Created    : 07-11-2025 10:06
'* Author     : VBATools
'* Copyright  : VBATools
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Public Sub showForm()
    frmTestClass.Show
End Sub

'Public Sub testVisibleProperty()
'    ' Test the Visible property functionality
'    Dim accordion As clsAccordion
'    Set accordion = New clsAccordion
'    Call accordion.Initialize(Frame1, False) ' Assuming we have a frame control
'
'    ' Add some test items
'    Call accordion.AddItem("Item 1", "Content of item 1")
'    Call accordion.AddItem("Item 2", "Content of item 2")
'    Call accordion.AddItem("Item 3", "Content of item 3")
'
'    ' Test hiding an item
'    accordion.item(2).Visible = False
'    Debug.Print "Item 2 visibility: " & accordion.item(2).Visible
'    Debug.Print "Total visible items should adjust positions accordingly"
'
'    ' Test showing the item again
'    accordion.item(2).Visible = True
'    Debug.Print "Item 2 visibility after showing: " & accordion.item(2).Visible
'    Debug.Print "Item 2 should return to its original position with proper layout adjustment"
'
'    ' Clean up
'    Set accordion = Nothing
'End Sub



