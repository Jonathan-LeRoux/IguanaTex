VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LogFileViewer 
   Caption         =   "Error in Latex Code"
   ClientHeight    =   6960
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8844.001
   OleObjectBlob   =   "LogFileViewer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LogFileViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    Me.Height = 372
    Me.Width = 451
    ShowAcceleratorTip Me.CloseLogButton
    ShowAcceleratorTip Me.CmdButtonExternalEditor
    #If Mac Then
        ResizeUserForm Me
    #End If
End Sub

Private Sub UserForm_Activate()
    #If Mac Then
        MacEnableAccelerators Me
    #End If
End Sub

Sub CloseLogButton_Click()
    Dim SelStartPos As Long
    SelStartPos = LatexForm.TextWindow1.SelStart
    Dim TempPath As String
    TempPath = CleanPath(LatexForm.TextBoxTempFolder.Text)
    
    LatexForm.TextWindow1.Text = ReadAll(TempPath & DefaultFilePrefix & ".tex")

    CloseLogButton.Caption = "Close"
    Unload LogFileViewer
    'If LatexForm.isFormModeless Then
    '    LatexForm.Hide
    '    LatexForm.Show vbModal
    'End If
    LatexForm.TextWindow1.SetFocus
    If SelStartPos < Len(LatexForm.TextWindow1.Text) Then
        LatexForm.TextWindow1.SelStart = SelStartPos
    End If
End Sub

Sub CmdButtonExternalEditor_Click()
    Dim TempPath As String
    TempPath = CleanPath(LatexForm.TextBoxTempFolder.Text)
    If Not IsPathWritable(TempPath) Then Exit Sub

    LogFileViewer.Caption = ShellEscape(GetEditorPath()) & " " & ShellEscape(TempPath & DefaultFilePrefix & ".tex")
    CloseLogButton.Caption = "Reload modified code"
    #If Mac Then
        AppleScriptTask "IguanaTex.scpt", "MacExecute", GetEditorPath() & " " & ShellEscape(TempPath & DefaultFilePrefix & ".tex")
    #Else
        Shell ShellEscape(GetEditorPath()) & " " & ShellEscape(TempPath & DefaultFilePrefix & ".tex"), vbNormalFocus
    #End If
End Sub

#If Mac Then

#Else
' Mousewheel functions

Private Sub TextBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal x As Single, ByVal Y As Single)
    If Not Me Is Nothing Then
        HookListBoxScroll Me, Me.TextBox1
    End If
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
        UnhookListBoxScroll
End Sub

#End If
