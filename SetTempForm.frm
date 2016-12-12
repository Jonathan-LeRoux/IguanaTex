VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetTempForm 
   Caption         =   "Settings"
   ClientHeight    =   4711
   ClientLeft      =   14
   ClientTop       =   329
   ClientWidth     =   6286
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
    
    ' Temp folder
    res = TextBox1.Text
    If Right(res, 1) <> "\" Then
        res = res & "\"
    End If
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Temp Dir", REG_SZ, CStr(res)
    
    ' UTF8
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "UseUTF8", REG_DWORD, BoolToInt(CheckBoxUTF8.Value)
    
    ' PDF2PNG
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "UsePDF", REG_DWORD, BoolToInt(CheckBoxPDF.Value)
        
    ' GS command
    res = TextBoxGS.Text
    If Left(res, 1) = """" Then res = Mid(res, 2, Len(res) - 1)
    If Right(res, 1) = """" Then res = Left(res, Len(res) - 1)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "GS Command", REG_SZ, CStr(res)
    
    ' Path to ImageMagick Convert
    res = TextBoxIMconv.Text
    If Left(res, 1) = """" Then res = Mid(res, 2, Len(res) - 1)
    If Right(res, 1) = """" Then res = Left(res, Len(res) - 1)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "IMconv", REG_SZ, CStr(res)
    
    ' Time Out Interval for Processes
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "TimeOutTime", REG_DWORD, CLng(val(TextBoxTimeOut.Text))
    
    ' LaTeX Engine
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LaTeXEngine", REG_SZ, CStr(ComboBoxEngine.Text)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LaTeXEngineID", REG_DWORD, ComboBoxEngine.ListIndex
    
    Unload SetTempForm
End Sub


Private Sub CheckBoxPDF_Click()

    If CheckBoxPDF.Value = True Then
        TextBoxGS.Enabled = True
        TextBoxIMconv.Enabled = True
        ComboBoxEngine.Enabled = True
    Else
        TextBoxGS.Enabled = False
        TextBoxIMconv.Enabled = False
        ComboBoxEngine.Enabled = False
    End If

End Sub

Private Sub Reset_Click()
    CheckBoxUTF8.Value = True
    
    CheckBoxPDF.Value = False
    
    TextBoxGS.Text = "C:\Program Files (x86)\gs\gs9.15\bin\gswin32c.exe"
    
    TextBoxIMconv.Text = "C:\Program Files\ImageMagick\convert.exe"
    
    TextBoxTimeOut.Text = "60"
    
    ComboBoxEngine.ListIndex = 0
    
End Sub

Private Sub UserForm_Initialize()
    Dim res As String
    RegPath = "Software\IguanaTex"
    res = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Temp Dir", "c:\temp")
    If Right(res, 1) <> "\" Then
        res = res & "\"
    End If
    TextBox1.Text = res
    
    CheckBoxUTF8.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "UseUTF8", True)
    
    CheckBoxPDF.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "UsePDF", False)
    
    TextBoxGS.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "GS Command", "C:\Program Files (x86)\gs\gs9.15\bin\gswin32c.exe")
    
    TextBoxIMconv.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "IMconv", "C:\Program Files\ImageMagick\convert.exe")
    
    TextBoxTimeOut.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TimeOutTime", "60")
    
    ComboBoxEngine.List = Array("pdflatex", "xelatex", "lualatex")
    'With ComboBoxEngine
    '    .AddItem "pdflatex"
    '    .AddItem "xelatex"
    '    .AddItem "lualatex"
    'End With
    ComboBoxEngine.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LaTeXEngineID", 0)
End Sub

Private Function BoolToInt(val) As Long
    If val Then
        BoolToInt = 1&
    Else
        BoolToInt = 0&
    End If
End Function
