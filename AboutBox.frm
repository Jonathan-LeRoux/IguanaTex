VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutBox 
   Caption         =   "IguanaTex"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   330
   ClientWidth     =   8820.001
   OleObjectBlob   =   "AboutBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseAboutButton_Click()
    Unload AboutBox
End Sub


Private Sub LabelURL_Click()
    Link = "http://www.jonathanleroux.org/software/iguanatex/"
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", Link)
        
    If (lSuccess = 0) Then
        MsgBox "Cannot open " & Link
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    
End Sub
