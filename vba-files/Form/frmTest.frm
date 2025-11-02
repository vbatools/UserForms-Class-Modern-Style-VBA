VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTest 
   Caption         =   "UserForm1"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9945.001
   OleObjectBlob   =   "frmTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MStyleItem      As clsModernStyle


Private Sub btnItem_Click()
    'MStyleItem.getItemByIndex(1)
    With MStyleItem.getItemByName(TextBox1.Name)
        .ColorBackGroundOff = vbRed
        .ColorBackGroundOn = vbGreen
    End With
End Sub

Private Sub btnStyle_Click()
    Call MStyleItem.Initialize(Me)
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .left = Application.left + 0.5 * (Application.width - .width)
        .top = Application.top + 0.5 * (Application.height - .height)
    End With

    ComboBox1.AddItem 1
    ComboBox1.AddItem 2
    ComboBox1.AddItem 3
    ComboBox1.AddItem 4

    ComboBox2.AddItem 1
    ComboBox2.AddItem 2
    ComboBox2.AddItem 3
    ComboBox2.AddItem 4
    
    ListBox1.AddItem 1
    ListBox1.AddItem 2
    ListBox1.AddItem 3


    Set MStyleItem = New clsModernStyle
End Sub
