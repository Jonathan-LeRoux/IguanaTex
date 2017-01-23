VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetTempForm 
   Caption         =   "Default Settings and Paths"
   ClientHeight    =   7230
   ClientLeft      =   15
   ClientTop       =   330
   ClientWidth     =   6285
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


Private Sub ButtonTeX2img_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker) 'msoFileDialogFolderPicker
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = TextBoxTeX2img.Text
    fd.Filters.Clear
    fd.Filters.Add "All Files", "*.*", 1
    
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            TextBoxTeX2img.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If

    Set fd = Nothing
    TextBoxTeX2img.SetFocus
End Sub

Private Sub ButtonTeXExePath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = TextBoxTeXExePath.Text
    
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            TextBoxTeXExePath.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If

    Set fd = Nothing
    TextBoxTeXExePath.SetFocus
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
    
    ' Vector or Bitmap (EMF or PNG)
    'SetRegistryValue HKEY_CURRENT_USER, RegPath, "EMFoutput", REG_DWORD, BoolToInt(CheckBoxEMF.Value)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "BitmapVector", REG_DWORD, ComboBoxBitmapVector.ListIndex
     
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
    
    ' Path to TeX2img (Vector output)
    res = TextBoxTeX2img.Text
    If Left(res, 1) = """" Then res = Mid(res, 2, Len(res) - 1)
    If Right(res, 1) = """" Then res = Left(res, Len(res) - 1)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "TeX2img Command", REG_SZ, CStr(res)
    
    ' Path to TeX Executables Folder
    res = TextBoxTeXExePath.Text
    If Left(res, 1) = """" Then res = Mid(res, 2, Len(res) - 1)
    If Right(res, 1) = """" Then res = Left(res, Len(res) - 1)
    If res <> "" And Right(res, 1) <> "\" Then res = res & "\"
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "TeXExePath", REG_SZ, CStr(res)
    
    ' Magic scaling factor to fine-tune the scaling of Vector displays
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "VectorScalingX", REG_SZ, TextBoxVectorScalingX.Text
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "VectorScalingY", REG_SZ, TextBoxVectorScalingY.Text
    
    ' Magic scaling factor to fine-tune the scaling of PNG displays
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "BitmapScalingX", REG_SZ, TextBoxBitmapScalingX.Text
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "BitmapScalingY", REG_SZ, TextBoxBitmapScalingY.Text
    
    ' Global dpi setting for latex output
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "OutputDpi", REG_DWORD, CLng(val(TextBoxDpi.Text))
    
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


Private Sub LabelDLgs_Click()
    Link = "http://www.ghostscript.com/download/gsdnld.html"
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", Link)
    If (lSuccess = 0) Then
        MsgBox "Cannot open " & Link
    End If
End Sub

Private Sub LabelDLImageMagick_Click()
    Link = "http://www.imagemagick.org/script/binary-releases.php"
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", Link)
    If (lSuccess = 0) Then
        MsgBox "Cannot open " & Link
    End If
End Sub

Private Sub LabelDLTeX2img_Click()
    Link = "http://www.math.sci.hokudai.ac.jp/~abenori/soft/bin/TeX2img_2.0.1.zip"
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", Link)
    If (lSuccess = 0) Then
        MsgBox "Cannot open " & Link
    End If
End Sub

Private Sub LabelTeX2imgGithub_Click()
    Link = "https://github.com/abenori/TeX2img"
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", Link)
    If (lSuccess = 0) Then
        MsgBox "Cannot open " & Link
    End If
End Sub

Private Sub LabelDLtexstudio_Click()
    Link = "http://www.texstudio.org/"
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", Link)
    If (lSuccess = 0) Then
        MsgBox "Cannot open " & Link
    End If
End Sub

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
    
    'CheckBoxEMF.Value = False
    ComboBoxBitmapVector.ListIndex = 0
    
    TextBoxGS.Text = "C:\Program Files (x86)\gs\gs9.15\bin\gswin32c.exe"
    
    TextBoxIMconv.Text = "C:\Program Files\ImageMagick\convert.exe"
    
    TextBoxTeX2img.Text = "%USERPROFILE%\Downloads\TeX2img\TeX2imgc.exe"
    
    TextBoxTeXExePath.Text = ""
    
    TextBoxDpi.Text = "1200"
    
    TextBoxVectorScalingX.Text = "1"
    TextBoxVectorScalingY.Text = "1"
    
    TextBoxBitmapScalingX.Text = "1"
    TextBoxBitmapScalingY.Text = "1"
    
    TextBoxTimeOut.Text = "60"
    
    TextBoxFontSize.Text = "10"
    
    ComboBoxEngine.ListIndex = 0
    
    'SetPDFdependencies
    SetAbsRelDependencies
    
End Sub



Private Sub UserForm_Initialize()
    
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    
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
    
    TextBoxDpi.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "OutputDpi", "1200")
    
    TextBoxTimeOut.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TimeOutTime", "60")
    
    TextBoxFontSize.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "EditorFontSize", "10")
    
    TextBoxVectorScalingX.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "VectorScalingX", "1")
    TextBoxVectorScalingY.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "VectorScalingY", "1")
    
    TextBoxBitmapScalingX.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "BitmapScalingX", "1")
    TextBoxBitmapScalingY.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "BitmapScalingY", "1")
    
    TextBoxExternalEditor.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Editor", "C:\Program Files (x86)\TeXstudio\texstudio.exe")
    
    TextBoxTeX2img.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TeX2img Command", "%USERPROFILE%\Downloads\TeX2img\TeX2imgc.exe")
    
    TextBoxTeXExePath.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TeXExePath", "")
    
    'CheckBoxEMF.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "EMFoutput", False)
    ComboBoxBitmapVector.List = Array("Bitmap", "Vector")
    ComboBoxBitmapVector.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "BitmapVector", 0)
    
    LaTexEngineDisplayList = Array("latex", "pdflatex", "xelatex", "lualatex", "platex")
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
