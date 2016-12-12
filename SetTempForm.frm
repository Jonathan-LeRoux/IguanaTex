VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetTempForm 
   Caption         =   "Settings"
   ClientHeight    =   5369
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
Dim LaTexEngineList As Variant
Dim LaTexEngineDisplayList As Variant
Dim UsePDFList As Variant

Private Sub ButtonAbsTempPath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker) 'msoFileDialogFilePicker
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = AbsPathTextBox.Text
    
    If fd.Show = -1 Then

        For Each vrtSelectedItem In fd.SelectedItems

            AbsPathTextBox.Text = vrtSelectedItem

        Next vrtSelectedItem
    End If

    Set fd = Nothing
End Sub

Private Sub ButtonCancelTemp_Click()
    Unload SetTempForm
End Sub

Private Sub ButtonEditorPath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker) 'msoFileDialogFolderPicker
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = TextBoxExternalEditor.Text
    fd.Filters.Clear
    fd.Filters.Add "All Files", "*.*", 1
    
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            TextBoxExternalEditor.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If

    Set fd = Nothing
    TextBoxExternalEditor.SetFocus
End Sub

Private Sub ButtonGSPath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker) 'msoFileDialogFolderPicker
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = TextBoxGS.Text
    fd.Filters.Clear
    fd.Filters.Add "All Files", "*.*", 1
    
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            TextBoxGS.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If

    Set fd = Nothing
    TextBoxGS.SetFocus
End Sub

Private Sub ButtonIMPath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker) 'msoFileDialogFolderPicker
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = TextBoxIMconv.Text
    fd.Filters.Clear
    fd.Filters.Add "All Files", "*.*", 1
    
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            TextBoxIMconv.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If

    Set fd = Nothing
    TextBoxIMconv.SetFocus
End Sub

Private Sub ButtonSetTemp_Click()
    Dim RegPath As String
    Dim res As String
    RegPath = "Software\IguanaTex"
    
    ' Temp folder
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "AbsOrRel", REG_DWORD, BoolToInt(AbsPathButton.Value)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Abs Temp Dir", REG_SZ, CStr(AbsPathTextBox.Text)
    If Left(RelPathTextBox.Text, 2) = ".\" Then
        RelPathTextBox.Text = Mid(RelPathTextBox.Text, 3, Len(RelPathTextBox.Text) - 2)
    End If
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Rel Temp Dir", REG_SZ, CStr(RelPathTextBox.Text)
    
    If AbsPathButton.Value = True Then
        res = AbsPathTextBox.Text
    Else
        res = ".\" & RelPathTextBox.Text
    End If
    If Right(res, 1) <> "\" Then
        res = res & "\"
    End If
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Temp Dir", REG_SZ, CStr(res)
    
    ' UTF8
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "UseUTF8", REG_DWORD, BoolToInt(CheckBoxUTF8.Value)
    
    ' PDF2PNG
    'SetRegistryValue HKEY_CURRENT_USER, RegPath, "UsePDF", REG_DWORD, BoolToInt(CheckBoxPDF.Value)
        
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
    
    ' Path to External Editor
    res = TextBoxExternalEditor.Text
    If Left(res, 1) = """" Then res = Mid(res, 2, Len(res) - 1)
    If Right(res, 1) = """" Then res = Left(res, Len(res) - 1)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Editor", REG_SZ, CStr(res)
    
    ' Time Out Interval for Processes
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "TimeOutTime", REG_DWORD, CLng(val(TextBoxTimeOut.Text))
    
    ' Font size for text in editor/template windows
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "EditorFontSize", REG_DWORD, CLng(val(TextBoxFontSize.Text))
    
    ' LaTeX Engine
    'SetRegistryValue HKEY_CURRENT_USER, RegPath, "LaTeXEngine", REG_SZ, CStr(ComboBoxEngine.Text)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LaTeXEngineID", REG_DWORD, ComboBoxEngine.ListIndex
    
    Unload SetTempForm
End Sub


Private Sub AbsPathButton_Click()
    AbsPathButton.Value = True
    SetAbsRelDependencies
End Sub

'Private Sub ComboBoxEngine_Change()
'    SetPDFdependencies
'End Sub


Private Sub RelPathButton_Click()
    AbsPathButton.Value = False
    SetAbsRelDependencies
End Sub

Private Sub SetAbsRelDependencies()
    RelPathButton.Value = Not AbsPathButton.Value
    AbsPathTextBox.Enabled = AbsPathButton.Value
    RelPathTextBox.Enabled = RelPathButton.Value
End Sub

'Private Sub CheckBoxPDF_Click()
'
'    If CheckBoxPDF.Value = True Then
'        TextBoxGS.Enabled = True
'        TextBoxIMconv.Enabled = True
'    Else
'        TextBoxGS.Enabled = False
'        TextBoxIMconv.Enabled = False
'    End If
'End Sub

Private Sub SetPDFdependencies()
    If UsePDFList(ComboBoxEngine.ListIndex) = True Then
        TextBoxGS.Enabled = True
        TextBoxIMconv.Enabled = True
    Else
        TextBoxGS.Enabled = False
        TextBoxIMconv.Enabled = False
    End If
End Sub

Private Sub Reset_Click()
    AbsPathButton.Value = True
    
    CheckBoxUTF8.Value = True
    
    'CheckBoxPDF.Value = False
    
    TextBoxGS.Text = "C:\Program Files (x86)\gs\gs9.15\bin\gswin32c.exe"
    
    TextBoxIMconv.Text = "C:\Program Files\ImageMagick\convert.exe"
    
    TextBoxTimeOut.Text = "60"
    
    TextBoxFontSize.Text = "10"
    
    ComboBoxEngine.ListIndex = 0
    
    'SetPDFdependencies
    SetAbsRelDependencies
    
End Sub

Private Sub UserForm_Initialize()
    Dim res As String
    RegPath = "Software\IguanaTex"
    
    res = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Abs Temp Dir", "c:\temp\")
    If Right(res, 1) <> "\" Then
        res = res & "\"
    End If
    AbsPathTextBox.Text = res
    
    res = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Rel Temp Dir", "")
    RelPathTextBox.Text = res
    
    AbsPathButton.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "AbsOrRel", True)
    
    CheckBoxUTF8.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "UseUTF8", True)
    
    TextBoxGS.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "GS Command", "C:\Program Files (x86)\gs\gs9.15\bin\gswin32c.exe")
    
    TextBoxIMconv.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "IMconv", "C:\Program Files\ImageMagick\convert.exe")
    
    TextBoxTimeOut.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TimeOutTime", "60")
    
    TextBoxFontSize.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "EditorFontSize", "10")
    
    TextBoxExternalEditor.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Editor", "C:\Program Files (x86)\TeXstudio\texstudio.exe")
    
    LaTexEngineDisplayList = Array("latex (DVI->PNG)", "pdflatex (PDF->PNG)", "xelatex (PDF->PNG)", "lualatex (PDF->PNG)", "platex (PDF->PNG)")
    UsePDFList = Array(False, True, True, True, True)
    
    ComboBoxEngine.List = LaTexEngineDisplayList
    ComboBoxEngine.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LaTeXEngineID", 0)
    'CheckBoxPDF.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "UsePDF", False)
    
    'SetPDFdependencies
    SetAbsRelDependencies
End Sub

Private Function BoolToInt(val) As Long
    If val Then
        BoolToInt = 1&
    Else
        BoolToInt = 0&
    End If
End Function
