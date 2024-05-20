VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "UserForm1"
   ClientHeight    =   6216
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8052
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














Private Sub CommandButton1_Click()
    VBAImg.OpenImg
End Sub
Private Sub CommandButton2_Click()
    VBAImg.SaveImg
End Sub
Private Sub Invert_Click()
    VBAImg.Inv
End Sub

Private Sub Label1_Click()

End Sub

Private Sub ScrollBar1_Change()

End Sub

Private Sub SpinButton1_Change()
    ComboBox1.ListWidth = SpinButton1.Value
    Label1.Caption = "ListWidth = " _
    & SpinButton1.Value
End Sub
 
Private Sub UserForm_Initialize()
    Dim i As Integer
    
    For i = 1 To 20
        ComboBox1.AddItem "Choice " _
        & (ComboBox1.ListCount + 1)
    Next i
    
    SpinButton1.Min = 0
    SpinButton1.Max = 255
    SpinButton1.Value = Val(ComboBox1.ListWidth)
    SpinButton1.SmallChange = 5
    Label1.Caption = "ListWidth = " _
    & SpinButton1.Value
End Sub

