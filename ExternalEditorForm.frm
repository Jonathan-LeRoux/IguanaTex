VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExternalEditorForm 
   Caption         =   "External Editor"
   ClientHeight    =   2688
   ClientLeft      =   84
   ClientTop       =   396
   ClientWidth     =   5232
   OleObjectBlob   =   "ExternalEditorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExternalEditorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    Me.Height = 158
    Me.Width = 270
    #If Mac Then
        ResizeUserForm Me
    #End If
End Sub

Private Sub UserForm_Activate()
    #If Mac Then
        MacEnableAccelerators Me
    #End If
End Sub

Sub CmdButtonCancel_Click()
    Unload ExternalEditorForm
End Sub

Private Sub LoadTextIntoLatexForm()
    Dim SelStartPos As Long
    SelStartPos = LatexForm.TextWindow1.SelStart

    Dim TempPath As String
    TempPath = CleanPath(LatexForm.TextBoxTempFolder.Text)
    
    LatexForm.TextWindow1.Text = ReadAll(TempPath & "ext_" & DefaultFilePrefix & ".tex")

    Unload ExternalEditorForm
    LatexForm.TextWindow1.SetFocus
    If SelStartPos < Len(LatexForm.TextWindow1.Text) Then
        LatexForm.TextWindow1.SelStart = SelStartPos
    End If

End Sub

Sub CmdButtonReload_Click()
    LoadTextIntoLatexForm
    LatexForm.Hide
    LatexForm.Show vbModal
End Sub

Sub CmdButtonGenerate_Click()
    LoadTextIntoLatexForm
    DoEvents
    LatexForm.ButtonRun_Click
End Sub

Public Sub LaunchExternalEditor(TempPath As String, LatexCode As String)
    ' Put the temporary path in the right format and test if it is writable
    TempPath = CleanPath(TempPath)
    If Not IsPathWritable(TempPath) Then Exit Sub
    
    Dim FilePrefix As String
    FilePrefix = "ext_" & DefaultFilePrefix
    
    ' Write latex to a temp file
    WriteToFile TempPath, FilePrefix, LatexCode
    
    ' Launch external editor
    On Error GoTo ShellError
    #If Mac Then
        AppleScriptTask "IguanaTex.scpt", "MacExecute", GetEditorPath() & " " & ShellEscape(TempPath & FilePrefix & ".tex")
    #Else
        Shell ShellEscape(GetEditorPath()) & " " & ShellEscape(TempPath & FilePrefix & ".tex"), vbNormalFocus
    #End If
    
    ' Show dialog form to reload from file or cancel
    Me.Show
    Exit Sub
    
ShellError:
    MsgBox "Error Launching External Editor." & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
    Exit Sub
End Sub



