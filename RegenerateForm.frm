VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegenerateForm 
   Caption         =   "Regenerating"
   ClientHeight    =   1926
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   4032
   OleObjectBlob   =   "RegenerateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegenerateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    Me.Height = 120
    Me.Width = 210
    #If Mac Then
        ResizeUserForm Me
    #End If
End Sub

Private Sub CommandButtonCancel_Click()
    RegenerateContinue = False
    Unload RegenerateForm
End Sub
