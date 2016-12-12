VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExternalEditorForm 
   Caption         =   "External Editor"
   ClientHeight    =   2037
   ClientLeft      =   21
   ClientTop       =   336
   ClientWidth     =   5684
   OleObjectBlob   =   "ExternalEditorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExternalEditorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
End Sub

Private Sub CmdButtonCancel_Click()
    Unload ExternalEditorForm
End Sub

Private Sub CmdButtonReload_Click()
    SelStartPos = LatexForm.TextBox1.SelStart
    TempPath = LatexForm.TextBoxTempFolder.Text
    
    If Left(TempPath, 1) = "." Then
        Dim sPath As String
        sPath = ActivePresentation.path
        If Len(sPath) > 0 Then
            If Right(sPath, 1) <> "\" Then
                sPath = sPath & "\"
            End If
            TempPath = sPath & TempPath
        Else
            MsgBox "You need to have saved your presentation once to use a relative path."
            Exit Sub
        End If
    End If
    
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (TempPath & GetFilePrefix() & ".tex")
    LatexForm.TextBox1.Text = objStream.ReadText()

    Unload ExternalEditorForm
    LatexForm.TextBox1.SetFocus
    If SelStartPos < Len(LatexForm.TextBox1.Text) Then
        LatexForm.TextBox1.SelStart = SelStartPos
    End If
End Sub
