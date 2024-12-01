VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutBox 
   Caption         =   "IguanaTex"
   ClientHeight    =   5436
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8748.001
   OleObjectBlob   =   "AboutBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub CloseAboutButton_Click()
    Unload AboutBox
End Sub


Private Sub LabelURL_Click()
    OpenURL "https://www.jonathanleroux.org/software/iguanatex/"
End Sub

Private Sub LabelGithub_Click()
    OpenURL "https://github.com/Jonathan-LeRoux/IguanaTex"
End Sub

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    Me.Height = 300
    Me.Width = 448
    Me.LabelAuthors.Caption = "by Jonathan Le Roux and Zvika Ben-Haim" & NEWLINE & NEWLINE & _
                              "Mac version by Tsung-Ju Chiang and Jonathan Le Roux"
    ShowAcceleratorTip Me.CloseAboutButton
    #If Mac Then
        ResizeUserForm Me
    #End If
End Sub

Private Sub UserForm_Activate()
    #If Mac Then
        MacEnableAccelerators Me
    #End If
End Sub

