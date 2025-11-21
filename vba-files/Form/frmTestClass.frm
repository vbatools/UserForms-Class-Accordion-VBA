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

Dim accord          As clsAccordion

Private Sub chbOpenAll_Change()
    Call accord.OpenAll(chbOpenAll.Value)
End Sub

Private Sub chbEnabled_Click()
    accord.item(2).Enabled = chbEnabled.Value
End Sub

Private Sub btnRemoveItem_Click()
    Debug.Print accord.RemoveItem(2)
End Sub

Private Sub btnShowItem_Click()
    ' Show the second item
    accord.item(2).Visible = True
End Sub

Private Sub btnGetItems_Click()
    Debug.Print accord.item(1).Name

    Dim coll        As Collection
    Set coll = accord.Items
End Sub

Private Sub btnTitleForeColor_Click()
    accord.item(2).TitleBtn.ForeColor = vbRed
End Sub

Private Sub btnColorItem_Click()
    Debug.Print accord.item(2).BackColor
    accord.item(2).BackColor = vbRed
End Sub

Private Sub btnForeColorItem_Click()
    Debug.Print accord.item(3).ForeColor
    accord.item(3).ForeColor = vbRed
End Sub

Private Sub chbVisible_Click()
    accord.item(2).Visible = chbVisible.Value
End Sub

Private Sub lbRemoveAll_Click()
    lbRemoveAll.Caption = accord.Count
    accord.RemoveAll
    Call accord.AddItem("Settings Panel", "Contains application settings and configuration options", False, 20, 210, vbBlue)
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .left = Application.left + 0.5 * (Application.width - .width)
        .Top = Application.Top + 0.5 * (Application.height - .height)
    End With
    Set accord = New clsAccordion
    Call accord.Initialize(Frame1, False)
    Call accord.AddItem("Introduction", "Overview of the ion control functionality")

    Call accord.AddItem("Configuration", "Application configuration settings", False, 20, 200)
    Call accord.AddItem("User Preferences", "User-specific preferences and options", False, 20, 200)
    Call accord.AddItem("Security Settings", "Security and access control options", False, 20, 200)
    Call accord.AddItem("Display Options", "Display and appearance settings", False, 20, 200)
    Call accord.AddItem("Advanced Settings", "Advanced configuration options", False, 20, 200)
    With accord.AddItem("Help & Support", "Documentation and support resources", False, 20, 50, vbBlue, 1, 2)
        .TitleLabel.Font.Bold = True
    End With

End Sub
