VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetTempForm 
   Caption         =   "Set Temporary Folder"
   ClientHeight    =   1659
   ClientLeft      =   21
   ClientTop       =   336
   ClientWidth     =   6272
   OleObjectBlob   =   "SetTempForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SetTempForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub ButtonCancelTemp_Click()
    
    Unload SetTempForm
End Sub

Private Sub ButtonSetTemp_Click()
    Dim RegPath As String
    Dim res As String
    RegPath = "Software\IguanaTex"
    res = TextBox1.Text
    If Right(res, 1) <> "\" Then
        res = res & "\"
    End If
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Temp Dir", REG_SZ, CStr(res)
    
    Unload SetTempForm
End Sub

Private Sub UserForm_Initialize()
    Dim res As String
    RegPath = "Software\IguanaTex"
    res = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Temp Dir", "c:\temp")
    If Right(res, 1) <> "\" Then
        res = res & "\"
    End If
    TextBox1.Text = res
End Sub
