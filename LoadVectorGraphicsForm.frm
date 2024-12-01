VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadVectorGraphicsForm 
   Caption         =   "Load Vector Graphics File"
   ClientHeight    =   3408
   ClientLeft      =   96
   ClientTop       =   324
   ClientWidth     =   7020
   OleObjectBlob   =   "LoadVectorGraphicsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadVectorGraphicsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBoxVectorOutputType_Change()
    SetVectorTypeDependencies
End Sub

Private Sub SetVectorTypeDependencies()
    If ComboBoxVectorOutputType.ListIndex = 0 Then
        CheckBoxCleanUp.value = False
        CheckBoxCleanUp.Enabled = False
        CheckBoxConvertLines.value = False
        CheckBoxConvertLines.Enabled = False
    Else
        CheckBoxCleanUp.Enabled = True
        CheckBoxConvertLines.Enabled = True
        CheckBoxCleanUp.value = GetITSetting("LoadVectorFileCleanUp", True)
        CheckBoxConvertLines.value = GetITSetting("LoadVectorFileConvertLines", False)
    End If
End Sub

Sub CommandButtonSave_Click()
    SetITSetting "LoadVectorFileConvertLines", REG_DWORD, BoolToInt(CheckBoxConvertLines.value)
    SetITSetting "LoadVectorFileScaling", REG_SZ, textboxScalor.Text
    SetITSetting "LoadVectorFileCalibrationX", REG_SZ, TextBoxCalibrationX.Text
    SetITSetting "LoadVectorFileCalibrationY", REG_SZ, TextBoxCalibrationY.Text
    SetITSetting "LoadVectorFileOutputTypeIdx", REG_DWORD, ComboBoxVectorOutputType.ListIndex
    SetITSetting "LoadVectorFileCleanUp", REG_DWORD, BoolToInt(CheckBoxCleanUp.value)
End Sub

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    Me.Height = 194
    Me.Width = 355
    #If Mac Then
        Me.LabelInsertPath.Caption = "Insert path of .pdf/.dvi/.xdv/.ps/.eps/.svg file:"
        ResizeUserForm Me
        ComboBoxVectorOutputType.Enabled = False
    #End If
    textboxScalor.Text = GetITSetting("LoadVectorFileScaling", "1")
    CheckBoxConvertLines.value = GetITSetting("LoadVectorFileConvertLines", False)
    TextBoxCalibrationX.Text = GetITSetting("LoadVectorFileCalibrationX", "1")
    TextBoxCalibrationY.Text = GetITSetting("LoadVectorFileCalibrationY", "1")
    ComboBoxVectorOutputType.List = Array("SVG via PDF w/ dvisvgm", "EMF w/ pdfiumdraw")
    ComboBoxVectorOutputType.ListIndex = GetITSetting("LoadVectorFileOutputTypeIdx", 0)
    CheckBoxCleanUp.value = GetITSetting("LoadVectorFileCleanUp", True)
    SetVectorTypeDependencies
    ShowAcceleratorTip Me.ButtonLoadFile
    ShowAcceleratorTip Me.ButtonCancel
    ShowAcceleratorTip Me.CommandButtonSave
    
End Sub

Private Sub UserForm_Activate()
    #If Mac Then
        MacEnableAccelerators Me
    #End If
End Sub

Sub ButtonCancel_Click()
    Unload LoadVectorGraphicsForm
End Sub

Private Function isInsertableVectorFile(file As String) As Boolean
    Dim Ext As String
    Ext = GetExtension(file)
    #If Mac Then
        isInsertableVectorFile = Ext = "pdf" Or Ext = "dvi" Or Ext = "xdv" Or Ext = "ps" Or Ext = "eps" Or Ext = "svg"
    #Else
        isInsertableVectorFile = Ext = "pdf" Or Ext = "dvi" Or Ext = "xdv" Or Ext = "ps" Or Ext = "eps" Or Ext = "emf" Or Ext = "svg"
    #End If
End Function

Sub ButtonPath_Click()
    #If Mac Then
        TextBoxFile.Text = MacChooseFileOfType("pdf,dvi,xdv,ps,eps,svg")
    #Else
        TextBoxFile.Text = BrowseFilePath(TextBoxFile.Text, "Vector graphics files", "*.pdf;*.dvi;*.xdv;*.ps;*.eps;*.emf;*.svg", "&Select file")
    #End If
    TextBoxFile.SetFocus
End Sub

Private Sub TextBoxFile_Change()
    Dim path As String, Ext As String
    path = TextBoxFile.Text
    Ext = GetExtension(path)
    ButtonLoadFile.Enabled = FileExists(path) And isInsertableVectorFile(path)
    If Ext = "emf" And isInsertableVectorFile(path) Then
        ComboBoxVectorOutputType.ListIndex = 1
        ComboBoxVectorOutputType.Enabled = False
        SetVectorTypeDependencies
    ElseIf Ext = "svg" Then
        ComboBoxVectorOutputType.ListIndex = 0
        ComboBoxVectorOutputType.Enabled = False
        SetVectorTypeDependencies
    End If
End Sub

Sub ButtonLoadFile_Click()
    DoInsertVectorGraphicsFile
    Unload LoadVectorGraphicsForm
End Sub


Private Sub DoInsertVectorGraphicsFile()
    Dim NewShape As Shape
    Dim TimeOutTimeString As String
    Dim TimeOutTime As Long
    TimeOutTimeString = GetITSetting("TimeOutTime", "20") ' Wait 20 seconds for the processes to complete
    TimeOutTime = val(TimeOutTimeString) * 1000
    Dim debugMode As Boolean
    debugMode = False
    #If Mac Then
        Dim fs As New MacFileSystemObject
    #Else
        Dim fs As New FileSystemObject
    #End If
    
    Dim StartFolder As String
    ' The StartFolder doesn't really matter here because everything is relative to the input file,
    ' we just need any folder from which to launch the commands
    'If ActivePresentation.path <> vbNullString Then
    '    StartFolder = ActivePresentation.path
    'Else
    If GetTempPath() <> vbNullString Then
        StartFolder = GetTempPath()
    Else
        #If Mac Then
            StartFolder = "/"
        #Else
            StartFolder = "C:\"
        #End If
    End If
    
    Dim posX As Single, posY As Single, ScalingX As Single, ScalingY As Single
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    If Sel.Type = ppSelectionShapes Then
    ' if something is selected on a slide, use its position for the new display
            posX = Sel.ShapeRange(1).Left
            posY = Sel.ShapeRange(1).Top
    Else
        posX = 200
        posY = 200
    End If
    ScalingX = textboxScalor.value * TextBoxCalibrationX.value
    ScalingY = textboxScalor.value * TextBoxCalibrationY.value
    
    ' Get the path and extension of the file to be inserted
    Dim path As String, Ext As String, pdfPath As String, psPath As String
    path = TextBoxFile.Text
    Ext = GetExtension(path)
    
    Dim TeXExePath As String, TeXExeExt As String
    TeXExePath = GetITSetting("TeXExePath", DEFAULT_TEX_EXE_PATH)
    TeXExeExt = vbNullString
    Dim VectorOutputTypeList As Variant
    VectorOutputTypeList = Array("dvisvgm", "pdfiumdraw")
    Dim VectorOutputType As String
    VectorOutputType = VectorOutputTypeList(ComboBoxVectorOutputType.ListIndex)
    
    Dim ConvertLines As Boolean
    ConvertLines = CheckBoxConvertLines.value
    Dim CleanUp As Boolean
    CleanUp = CheckBoxCleanUp.value
    
    Dim RetVal As Long
    Dim ErrorMessage As String
    Dim RunCommand As String
    
    If Ext = "svg" Or Ext = "emf" Then
        Set NewShape = AddDisplayShape(path, posX, posY)
    Else
        If VectorOutputType = "dvisvgm" Then ' Convert to SVG
            Dim libgsPath As String
            libgsPath = GetITSetting("Libgs", DEFAULT_LIBGS)
            Dim libgsString As String
            If libgsPath <> vbNullString Then
                libgsString = " --libgs=" & ShellEscape(libgsPath)
            Else
                libgsString = vbNullString
            End If
            Dim InputTypeSwitch As String
            If Ext = "ps" Or Ext = "eps" Then
                InputTypeSwitch = " --eps"
            ElseIf Ext = "pdf" Then
                InputTypeSwitch = " --pdf"
            Else
                InputTypeSwitch = vbNullString
            End If
            Dim svgPath As String
            svgPath = path & "_tmp.svg"
            If fs.FileExists(svgPath) Then fs.DeleteFile svgPath
            RunCommand = ShellEscape(TeXExePath & "dvisvgm" & TeXExeExt) & InputTypeSwitch & " -o " & ShellEscape(svgPath) _
                                    & libgsString & " " & ShellEscape(path)
            RetVal& = Execute(RunCommand, StartFolder, debugMode, TimeOutTime)
            If (RetVal& <> 0 Or Not fs.FileExists(svgPath)) Then
                ' Error in EPS/PS/PDF to SVG conversion
                ErrorMessage = "Error while using dvisvgm to convert input file to SVG."
                ShowError ErrorMessage, RunCommand
                Exit Sub
            End If
            Ext = "svg"
            Set NewShape = AddDisplayShape(svgPath, posX, posY)
            If fs.FileExists(svgPath) Then fs.DeleteFile svgPath
        Else 'Convert to EMF
            If TeXExePath <> vbNullString Then TeXExeExt = ".exe"
            Dim DeleteTmpPDF As Boolean
            DeleteTmpPDF = False
            ' If .dvi/.xdv/.ps/.eps file, convert to .pdf first, using ps2pdf/eps2pdf/dvipdfmx
            If Ext = "ps" Or Ext = "eps" Or Ext = "dvi" Or Ext = "xdv" Then
                psPath = path
                pdfPath = path + "_tmp.pdf"
                If fs.FileExists(pdfPath) Then fs.DeleteFile pdfPath
                Dim pspdf_command As String
                Dim pdfpath_prefix As String
                pdfpath_prefix = vbNullString
                If Ext = "ps" Then
                    pspdf_command = "ps2pdf"
                ElseIf Ext = "eps" Then
                    pspdf_command = "epspdf"
                Else
                    pspdf_command = "dvipdfmx"
                    pdfpath_prefix = "-o "
                End If
                RunCommand = ShellEscape(TeXExePath & pspdf_command & TeXExeExt) & " " + ShellEscape(psPath) + " " + pdfpath_prefix + ShellEscape(pdfPath)
                RetVal& = Execute(RunCommand, StartFolder, debugMode, TimeOutTime)
                If (RetVal& <> 0 Or Not fs.FileExists(pdfPath)) Then
                    ErrorMessage = "DVI/XDV/PS/EPS to PDF conversion failed" _
                        & vbNewLine & "Make sure " & pspdf_command & " is installed (it comes with, e.g., Tex Live, MikTeX or Ghostscript) and can be run from anywhere via the command line"
                    ShowError ErrorMessage, RunCommand
                    Exit Sub
                End If
                Ext = "pdf"
                path = pdfPath
                DeleteTmpPDF = True
            End If
            ' Now we're dealing with a .pdf file
            
            ' Convert .pdf file to .emf using pdfiumdraw, which is part of TeX2img
            If Ext = "pdf" Then
                Dim emfPath As String
                emfPath = path + "_tmp.emf"
                If fs.FileExists(emfPath) Then fs.DeleteFile emfPath
                Dim tex2img_command As String
                Dim pdfiumdraw_command As String
                tex2img_command = GetITSetting("TeX2img Command", DEFAULT_TEX2IMG_COMMAND)
                pdfiumdraw_command = GetFolderFromPath(tex2img_command) & "pdfiumdraw.exe"
                RunCommand = ShellEscape(pdfiumdraw_command) & " --extent=50 --emf --transparent --pages=1 --output=" & ShellEscape(emfPath) _
                                    & " " & ShellEscape(path)
                RetVal& = Execute(RunCommand, StartFolder, debugMode, TimeOutTime)
                If (RetVal& <> 0 Or Not fs.FileExists(emfPath)) Then
                    ErrorMessage = " PDF to EMF conversion failed" _
                        & vbNewLine & "Make sure to correctly set the path to Tex2imgc.exe in Main Settings." _
                        & vbNewLine & "IguanaTex uses that path to find pdfiumdraw.exe."
                    ShowError ErrorMessage, RunCommand
                    Exit Sub
                End If
                Ext = "emf"
                Set NewShape = AddDisplayShape(emfPath, posX, posY)
                If Not debugMode Then
                    If fs.FileExists(emfPath) Then fs.DeleteFile emfPath
                    If DeleteTmpPDF Then
                        If fs.FileExists(pdfPath) Then fs.DeleteFile pdfPath
                    End If
                End If
            End If
        End If
    End If
    
    If Ext = "emf" Then
        Set NewShape = ConvertEMF(NewShape, ScalingX, ScalingY, posX, posY, Ext, ConvertLines, CleanUp)
    ElseIf Ext = "svg" Then
        Set NewShape = convertSVG(NewShape, ScalingX, ScalingY, posX, posY)
    Else
        MsgBox "We got lost somehow."
    End If
    NewShape.Select
End Sub


