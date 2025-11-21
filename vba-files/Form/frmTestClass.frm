VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestClass 
   Caption         =   "UserForm1"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12480
   OleObjectBlob   =   "frmTestClass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim acord           As clsAcardion

Private Sub chbOpenAll_Change()
    Call acord.OpenAll(chbOpenAll.Value)
End Sub

Private Sub chbEnabled_Click()
    acord.item(2).Enabled = chbEnabled.Value
End Sub

Private Sub btnRemoveItem_Click()
    Debug.Print acord.RemoveItem(2)
End Sub

Private Sub btnGetItems_Click()
    Debug.Print acord.item(1).Name

    Dim coll        As Collection
    Set coll = acord.Items
End Sub

Private Sub btnTitleForeColor_Click()
    acord.item(2).TitleBtn.ForeColor = vbRed
End Sub

Private Sub btnColorItem_Click()
    Debug.Print acord.item(2).BackColor
    acord.item(2).BackColor = vbRed
End Sub

Private Sub btnForeColorItem_Click()
    Debug.Print acord.item(3).ForeColor
    acord.item(3).ForeColor = vbRed
End Sub

Private Sub lbRemoveAll_Click()
    lbRemoveAll.Caption = acord.Count
    acord.RemoveAll
    Call acord.AddItem("Level 5 text", "Level 1 text text", False, 20, 210, vbBlue)
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .left = Application.left + 0.5 * (Application.width - .width)
        .Top = Application.Top + 0.5 * (Application.height - .height)
    End With
    Set acord = New clsAcardion
    Call acord.Initialize(Frame1, False)
    Call acord.AddItem("Level 1 text", "Level  1 text text")

    Call acord.AddItem("Level 2 text", "Level 2 text text", False, 20, 200)
    Call acord.AddItem("Level 21 text", "Level 2 text text", False, 20, 200)
    Call acord.AddItem("Level 22 text", "Level 2 text text", False, 20, 200)
    Call acord.AddItem("Level 23 text", "Level 2 text text", False, 20, 200)
    Call acord.AddItem("Level 24 text", "Level 2 text text", False, 20, 200)
    With acord.AddItem("Level 3 text", "Level 3 text text text text text text", False, 20, 50, vbBlue, 1, 2)
        .TitleLabel.Font.Bold = True
    End With

End Sub
