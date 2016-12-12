VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BatchEditForm 
   Caption         =   "Batch edit"
   ClientHeight    =   5474
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   4711
   OleObjectBlob   =   "BatchEditForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BatchEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RegPath As String
Dim LaTexEngineDisplayList As Variant


Private Sub UserForm_Initialize()
    LoadSettings
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    
End Sub

Private Sub LoadSettings()
    RegPath = "Software\IguanaTex"
    LaTexEngineDisplayList = Array("latex (DVI)", "pdflatex", "xelatex", "lualatex", "platex")
    ComboBoxLaTexEngine.List = LaTexEngineDisplayList
    ComboBoxLaTexEngine.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LaTeXEngineID", 0)
    TextBoxTempFolder.Text = GetTempPath()
    'CheckBoxEMF.Value = CBool(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "EMFoutput", False))
    ComboBoxBitmapVector.List = Array("Bitmap", "Vector")
    ComboBoxBitmapVector.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "BitmapVector", 0)
    
    TextBoxLocalDPI.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "OutputDpi", "1200")
    textboxSize.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "PointSize", "20")
    checkboxTransp.Value = CBool(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Transparent", True))
    CheckBoxResetFormat.Value = False
    
    CheckBoxModifyEngine.Value = False
    CheckBoxModifyTempFolder.Value = False
    CheckBoxModifyBitmapVector.Value = False
    CheckBoxModifyLocalDPI.Value = False
    CheckBoxModifySize.Value = False
    CheckBoxModifyTransparency.Value = False
    CheckBoxModifyResetFormat.Value = False
    CheckBoxReplace.Value = False
    Apply_CheckBoxModifyEngine
    Apply_CheckBoxModifyTempFolder
    Apply_CheckBoxModifyBitmapVector
    Apply_CheckBoxModifyLocalDPI
    Apply_CheckBoxModifySize
    Apply_CheckBoxModifyTransparency
    Apply_CheckBoxModifyResetFormat
    Apply_CheckBoxReplace
End Sub

Sub ButtonRun_Click()
    BatchEditForm.Hide
    
    Call RegenerateSelectedDisplays
    
    Unload BatchEditForm
End Sub

Private Sub ButtonCancel_Click()
    Unload BatchEditForm
End Sub


' Enable/Disable Modifications
Private Sub CheckBoxModifyEngine_Click()
    Apply_CheckBoxModifyEngine
End Sub

Private Sub CheckBoxModifyTempFolder_Click()
    Apply_CheckBoxModifyTempFolder
End Sub

Private Sub CheckBoxModifyBitmapVector_Click()
    Apply_CheckBoxModifyBitmapVector
End Sub

Private Sub CheckBoxModifyLocalDPI_Click()
    Apply_CheckBoxModifyLocalDPI
End Sub

Private Sub CheckBoxModifySize_Click()
    Apply_CheckBoxModifySize
End Sub

Private Sub CheckBoxModifyTransparency_Click()
    Apply_CheckBoxModifyTransparency
End Sub

Private Sub CheckBoxModifyResetFormat_Click()
    Apply_CheckBoxModifyResetFormat
End Sub

Private Sub CheckBoxReplace_Click()
    Apply_CheckBoxReplace
End Sub

Private Sub Apply_CheckBoxModifyEngine()
    LabelEngine.Enabled = CheckBoxModifyEngine.Value
    ComboBoxLaTexEngine.Enabled = CheckBoxModifyEngine.Value
End Sub

Private Sub Apply_CheckBoxModifyTempFolder()
    LabelTempFolder.Enabled = CheckBoxModifyTempFolder.Value
    TextBoxTempFolder.Enabled = CheckBoxModifyTempFolder.Value
End Sub

Private Sub Apply_CheckBoxModifyBitmapVector()
    LabelOutput.Enabled = CheckBoxModifyBitmapVector.Value
    ComboBoxBitmapVector.Enabled = CheckBoxModifyBitmapVector.Value
End Sub

Private Sub Apply_CheckBoxModifyLocalDPI()
    LabelLocalDPI.Enabled = CheckBoxModifyLocalDPI.Value
    TextBoxLocalDPI.Enabled = CheckBoxModifyLocalDPI.Value
    LabelDPI.Enabled = CheckBoxModifyLocalDPI.Value
End Sub

Private Sub Apply_CheckBoxModifySize()
    LabelSize.Enabled = CheckBoxModifySize.Value
    textboxSize.Enabled = CheckBoxModifySize.Value
    LabelPTS.Enabled = CheckBoxModifySize.Value
End Sub

Private Sub Apply_CheckBoxModifyTransparency()
    checkboxTransp.Enabled = CheckBoxModifyTransparency.Value
End Sub

Private Sub Apply_CheckBoxModifyResetFormat()
    CheckBoxResetFormat.Enabled = CheckBoxModifyResetFormat.Value
End Sub


Private Sub Apply_CheckBoxReplace()
    LabelReplace.Enabled = CheckBoxReplace.Value
    TextBoxFind.Enabled = CheckBoxReplace.Value
    LabelWith.Enabled = CheckBoxReplace.Value
    TextBoxReplacement.Enabled = CheckBoxReplace.Value
End Sub


Private Sub ComboBoxBitmapVector_Change()
    Apply_BitmapVector_Change
End Sub

Private Sub Apply_BitmapVector_Change()
    If ComboBoxBitmapVector.ListIndex = 1 Then
        CheckBoxModifyLocalDPI.Value = False
        CheckBoxModifyTransparency.Value = False
        checkboxTransp.Value = True
    End If
    Apply_CheckBoxModifyLocalDPI
    Apply_CheckBoxModifyTransparency
End Sub



