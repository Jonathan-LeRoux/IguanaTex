VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BatchEditForm 
   Caption         =   "Batch edit"
   ClientHeight    =   5964
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   4956
   OleObjectBlob   =   "BatchEditForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BatchEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    LoadSettings
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    Me.Height = 322
    Me.Width = 256
    #If Mac Then
        ResizeUserForm Me
    #End If
End Sub

Private Sub UserForm_Activate()
    #If Mac Then
        MacEnableAccelerators Me
    #End If
End Sub

Private Sub LoadSettings()
    ComboBoxLaTexEngine.List = GetLaTexEngineDisplayList()
    ComboBoxLaTexEngine.ListIndex = GetITSetting("LaTeXEngineID", 0)
    TextBoxTempFolder.Text = GetTempPath()
    'CheckBoxEMF.Value = CBool(GetITSetting("EMFoutput", False))
    ComboBoxBitmapVector.List = GetBitmapVectorList()
    ComboBoxBitmapVector.ListIndex = GetITSetting("BitmapVector", 0)
    
    TextBoxLocalDPI.Text = GetITSetting("OutputDpi", "1200")
    textboxSize.Text = GetITSetting("PointSize", "20")
    checkboxTransp.value = CBool(GetITSetting("Transparent", True))
    CheckBoxResetFormat.value = False
    
    CheckBoxModifyEngine.value = False
    CheckBoxModifyTempFolder.value = False
    CheckBoxModifyBitmapVector.value = False
    CheckBoxModifyLocalDPI.value = False
    CheckBoxModifySize.value = False
    CheckBoxModifyPreserveSize.value = False
    CheckBoxModifyTransparency.value = False
    CheckBoxModifyResetFormat.value = False
    CheckBoxReplace.value = False
    Apply_CheckBoxModifyEngine
    Apply_CheckBoxModifyTempFolder
    Apply_CheckBoxModifyBitmapVector
    Apply_CheckBoxModifyLocalDPI
    Apply_CheckBoxModifySize
    Apply_CheckBoxModifyPreserveSize
    Apply_CheckBoxModifyTransparency
    Apply_CheckBoxModifyResetFormat
    Apply_CheckBoxReplace
End Sub

Public Sub ButtonRun_Click()
    BatchEditForm.Hide
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection

    RegenerateSelectedDisplays Sel
    
    Unload BatchEditForm
End Sub

Public Sub ButtonCancel_Click()
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

Private Sub CheckBoxModifyPreserveSize_Click()
    Apply_CheckBoxModifyPreserveSize
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
    LabelEngine.Enabled = CheckBoxModifyEngine.value
    ComboBoxLaTexEngine.Enabled = CheckBoxModifyEngine.value
End Sub

Private Sub Apply_CheckBoxModifyTempFolder()
    LabelTempFolder.Enabled = CheckBoxModifyTempFolder.value
    TextBoxTempFolder.Enabled = CheckBoxModifyTempFolder.value
End Sub

Private Sub Apply_CheckBoxModifyBitmapVector()
    LabelOutput.Enabled = CheckBoxModifyBitmapVector.value
    ComboBoxBitmapVector.Enabled = CheckBoxModifyBitmapVector.value
End Sub

Private Sub Apply_CheckBoxModifyLocalDPI()
    LabelLocalDPI.Enabled = CheckBoxModifyLocalDPI.value
    TextBoxLocalDPI.Enabled = CheckBoxModifyLocalDPI.value
    LabelDPI.Enabled = CheckBoxModifyLocalDPI.value
End Sub

Private Sub Apply_CheckBoxModifySize()
    LabelSize.Enabled = CheckBoxModifySize.value
    textboxSize.Enabled = CheckBoxModifySize.value
    LabelPTS.Enabled = CheckBoxModifySize.value
End Sub

Private Sub Apply_CheckBoxModifyPreserveSize()
    CheckBoxForcePreserveSize.Enabled = CheckBoxModifyPreserveSize.value
End Sub

Private Sub Apply_CheckBoxModifyTransparency()
    checkboxTransp.Enabled = CheckBoxModifyTransparency.value
End Sub

Private Sub Apply_CheckBoxModifyResetFormat()
    CheckBoxResetFormat.Enabled = CheckBoxModifyResetFormat.value
End Sub

Private Sub Apply_CheckBoxReplace()
    LabelReplace.Enabled = CheckBoxReplace.value
    TextBoxFind.Enabled = CheckBoxReplace.value
    LabelWith.Enabled = CheckBoxReplace.value
    TextBoxReplacement.Enabled = CheckBoxReplace.value
End Sub


Private Sub ComboBoxBitmapVector_Change()
    Apply_BitmapVector_Change
End Sub

Private Sub Apply_BitmapVector_Change()
    If ComboBoxBitmapVector.ListIndex = 1 Then
        CheckBoxModifyLocalDPI.value = False
        CheckBoxModifyTransparency.value = False
        CheckBoxModifyLocalDPI.Enabled = False
        CheckBoxModifyTransparency.Enabled = False
        checkboxTransp.value = True
    Else
        CheckBoxModifyLocalDPI.Enabled = True
        CheckBoxModifyTransparency.Enabled = False
    End If
    Apply_CheckBoxModifyLocalDPI
    Apply_CheckBoxModifyTransparency
End Sub

Private Sub CheckBoxForcePreserveSize_Click()
    If CheckBoxForcePreserveSize.value = True Then
        CheckBoxModifySize.Enabled = False
        CheckBoxModifySize.value = False
    Else
        CheckBoxModifySize.Enabled = True
    End If
    Apply_CheckBoxModifySize
End Sub

