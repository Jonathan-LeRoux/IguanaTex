VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrorForm 
   Caption         =   "Error while running process"
   ClientHeight    =   1848
   ClientLeft      =   156
   ClientTop       =   612
   ClientWidth     =   10062
   OleObjectBlob   =   "ErrorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ErrorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
'    Me.Height = 180
'    Me.Width = 344
'    #If Mac Then
'        ResizeUserForm Me
'    #End If
End Sub

Private Sub UserForm_Activate()
    Dim spacing As Long
    spacing = 8

    Me.Height = 180
    Me.Width = 344
    
    Dim LabelCommand As String
    Dim ErrorMessage As String
    LabelCommand = Me.LabelCommand.Caption
    Me.LabelCommand.Caption = vbNullString
    ErrorMessage = Me.LabelError.Caption
    Me.LabelError.Caption = vbNullString
    
    With Me.LabelError
        .AutoSize = True
        .WordWrap = True
        .Width = 324
        Me.LabelError.Caption = ErrorMessage
        .AutoSize = False
        .Height = .Height + 2
        .Width = 324
    End With
    With Me.LabelCommand
        .AutoSize = True
        .WordWrap = True
        .Width = 252
        Me.LabelCommand.Caption = LabelCommand
        .AutoSize = False
        .Height = .Height + 2
        .Width = 252
    End With
    
    If Me.LabelError.Caption = vbNullString Then
        Me.LabelError.Height = 0
        Me.LabelError.Top = 0
    End If
    
    Me.LabelLastCommandPrompt.Top = Me.LabelError.Top + Me.LabelError.Height + spacing
    Me.LabelCommand.Top = Me.LabelLastCommandPrompt.Top + Me.LabelLastCommandPrompt.Height + spacing / 2
    Me.CopyCommandButton.Top = Me.LabelCommand.Top + (Me.LabelCommand.Height - 2) / 2 - Me.CopyCommandButton.Height / 2
    Me.CloseErrorButton.Top = Me.LabelCommand.Top + Me.LabelCommand.Height + spacing
    Me.Height = Me.CloseErrorButton.Top + Me.CloseErrorButton.Height + spacing + 22
    
    #If Mac Then
        ResizeUserForm Me
    #End If
End Sub

Private Sub CloseErrorButton_Click()
    Me.Hide
End Sub

Private Sub CopyCommandButton_Click()
    Clipboard Me.LabelCommand.Caption
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
