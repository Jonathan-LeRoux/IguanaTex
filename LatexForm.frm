VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LatexForm 
   Caption         =   "IguanaTex"
   ClientHeight    =   5880
   ClientLeft      =   -288
   ClientTop       =   -1044
   ClientWidth     =   7668
   OleObjectBlob   =   "LatexForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LatexForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        
Private LaTexEngineList As Variant
Private LaTexDVIOptionsList As Variant
Private LatexmkPDFOptionsList As Variant
Private LatexmkDVIOptionsList As Variant
Private UsePDFList As Variant
Private UseDVIList As Variant

'Dim NumberOfTemplates As Long
Private TemplateSortedListString As String
Private TemplateSortedList() As String
Private TemplateNameSortedListString As String

Private FormHeightWidthSet As Boolean
Private DoneWithActivation As Boolean

Private theAppEventHandler As New AppEventHandler

#If Mac Then
    Public TextWindow1 As New TextWindow
    Public TextWindowTemplateCode As New TextWindow
#Else
    Public TextWindow1 As MSForms.TextBox
    Public TextWindowTemplateCode As MSForms.TextBox
#End If

Sub InitializeApp()
    Set theAppEventHandler.App = Application
    
    AddMenuItem "New Latex display...", "NewLatexEquation", 18 '226
    AddMenuItem "Edit Latex display...", "EditLatexEquation", 37
    AddMenuItem "Regenerate selection...", "RegenerateSelection", 19
    AddMenuItem "Convert to Shape...", "ConvertToVector", 153
    AddMenuItem "Convert to Picture...", "ConvertToBitmap", 931
    AddMenuItem "Settings...", "LoadSettingsForm", 548
    AddMenuItem "Insert vector file...", "InsertVectorGraphicsFile", 23
    
End Sub

Sub AddMenuItem(ByVal itemText As String, ByVal itemCommand As String, ByVal itemFaceId As Long)
    ' Check if we have already added the menu item
    Dim initialized As Boolean
    Dim bef As Integer
    initialized = False
    bef = 1
    Dim Menu As CommandBars
    Set Menu = Application.CommandBars
    Dim i As Long
    For i = 1 To Menu("Insert").Controls.count
        With Menu("Insert").Controls(i)
            If .Caption = itemText Then
                initialized = True
                Exit For
            ElseIf InStr(.Caption, "Dia&gram") Then
                bef = i
            End If
        End With
    Next
    
    ' Create the menu choice.
    If Not initialized Then
        Dim NewControl As CommandBarControl
        Set NewControl = Menu("Insert").Controls.Add _
                              (Type:=msoControlButton, _
                               before:=bef, _
                               Id:=itemFaceId)
        NewControl.Caption = itemText
        NewControl.OnAction = itemCommand
        NewControl.Style = msoButtonIconAndCaption
    End If
End Sub

Sub UnInitializeApp()
    
    RemoveMenuItem "New Latex display..."
    RemoveMenuItem "Edit Latex display..."
    RemoveMenuItem "Regenerate selection..."
    RemoveMenuItem "Convert to Shape..."
    RemoveMenuItem "Convert to Picture..."
    RemoveMenuItem "Settings..."
    RemoveMenuItem "Insert vector file..."
    ' Clean up older versions
    RemoveMenuItem "Regenerate selected displays..."
    RemoveMenuItem "Convert to EMF..."
    RemoveMenuItem "Convert to PNG..."

    
End Sub

Sub RemoveMenuItem(ByVal itemText As String)
    Dim Menu As CommandBars
    Set Menu = Application.CommandBars
    Dim i As Variant
    For i = 1 To Menu("Insert").Controls.count
        If Menu("Insert").Controls(i).Caption = itemText Then
            Menu("Insert").Controls(i).Delete
            Exit For
        End If
    Next
    

End Sub

Sub ButtonCancel_Click()
    Unload LatexForm
End Sub

Sub ButtonRun_Click()
    Dim TempPath As String
    TempPath = CleanPath(TextBoxTempFolder.Text)
    If Not IsPathWritable(TempPath) Then Exit Sub
    
    Dim FilePrefix As String
    FilePrefix = DefaultFilePrefix
    
    Dim debugMode As Boolean
    debugMode = checkboxDebug.value
    
    ' Check if an external editor is being used as default, if so, do not delete temp files to avoid issues with external editor
    Dim UseExternalEditor As Boolean
    UseExternalEditor = GetITSetting("UseExternalEditor", False)
    
    ' Read settings
    Dim LATEXENGINEID As Integer
    LATEXENGINEID = ComboBoxLaTexEngine.ListIndex
    Dim latex_command As String
    latex_command = LaTexEngineList(LATEXENGINEID)
    Dim latex_dvi_options As String
    latex_dvi_options = LaTexDVIOptionsList(LATEXENGINEID)
    Dim gs_command As String
    gs_command = GetITSetting("GS Command", DEFAULT_GS_COMMAND)
    Dim IMconv As String
    IMconv = GetITSetting("IMconv", DEFAULT_IM_CONV)
    Dim tex2img_command As String
    tex2img_command = GetITSetting("TeX2img Command", DEFAULT_TEX2IMG_COMMAND)
    Dim pdfiumdraw_command As String
    pdfiumdraw_command = GetFolderFromPath(tex2img_command) & "pdfiumdraw.exe"
    Dim UseLatexmk As Boolean
    UseLatexmk = GetITSetting("UseLatexmk", False)
    Dim latexmk_command As String
    latexmk_command = "latexmk"
    Dim latexmk_pdf_options As String
    latexmk_pdf_options = LatexmkPDFOptionsList(LATEXENGINEID)
    Dim latexmk_dvi_options As String
    latexmk_dvi_options = LatexmkDVIOptionsList(LATEXENGINEID)
    Dim AddAltText As Boolean
    AddAltText = GetITSetting("AddAltText", False)
    
    Dim TeXExePath As String, TeXExeExt As String
    TeXExePath = GetITSetting("TeXExePath", DEFAULT_TEX_EXE_PATH)
    TeXExeExt = vbNullString
    ' This does not seem to be necessary, at least not on my system,
    ' and it can break installations that use linux subsystems.
    '#If Mac Then
    '    ' no need to do anything for TeXExeExt on Mac
    '#Else
    '    If TeXExePath <> vbNullString Then TeXExeExt = ".exe"
    '#End If
    Dim libgsPath As String
    libgsPath = GetITSetting("Libgs", DEFAULT_LIBGS)
    Dim libgsString As String
    If libgsPath <> vbNullString Then
        libgsString = " --libgs=" & ShellEscape(libgsPath)
    Else
        libgsString = vbNullString
    End If
        
    Dim UseDVI As Boolean
    UseDVI = UseDVIList(LATEXENGINEID)
    Dim UsePDF As Boolean
    UsePDF = UsePDFList(LATEXENGINEID)
    
    Dim UseVector As Boolean
    Dim BitmapVector As Integer
    BitmapVector = ComboBoxBitmapVector.ListIndex
    UseVector = Not (BitmapVector = 0)
    Dim VectorOutputType As String
    VectorOutputType = GetITSetting("VectorOutputType", DEFAULT_VECTOR_OUTPUT_TYPE)
    Dim PictureOutputType As String
    PictureOutputType = GetITSetting("PictureOutputType", DEFAULT_PICTURE_OUTPUT_TYPE)
    
    Dim OutputType As String
    Dim OutputExt As String
    
    If UseVector Then
        If VectorOutputType = "dvisvgm" Then
            ' "dvisvgm via DVI" only needs DVI/XDV, no PDF, whatever the engine
            UseDVI = True
            UsePDF = False
        Else
            ' "pdfiumdraw" and "dvisvgm via PDF" both require PDF, whatever the engine.
            ' The last option, "tex2img", does not care how this is set.
            UsePDF = True
        End If
    Else
        #If Mac Then
            ' For PNG on Mac, we force the use of DVI as it's a pain to convert PDF to PNG with proper DPI
            If PictureOutputType = "PNG" Then
                UseDVI = True
                UsePDF = False
            End If
        #End If
    End If
    
    Dim TimeOutTimeString As String
    Dim TimeOutTime As Long
    TimeOutTimeString = GetITSetting("TimeOutTime", "20") ' Wait N seconds for the processes to complete
    TimeOutTime = val(TimeOutTimeString) * 1000
    
    Dim OutputDpiString As String
    OutputDpiString = TextBoxLocalDPI.Text
    Dim OutputDpi As Long
    OutputDpi = val(OutputDpiString)
    
    ' Read current dpi in: this will be used when rescaling
    Dim dpi As Double, default_screen_dpi As Double
    dpi = 96 'lDotsPerInch ' I'm not convinced that this is the right thing to do, so for now I stop trying to take dpi into account
    default_screen_dpi = 96
    Dim VectorScalingX As Single, VectorScalingY As Single, BitmapScalingX As Single, BitmapScalingY As Single
    VectorScalingX = dpi / default_screen_dpi * val(GetITSetting("VectorScalingX", "1"))
    VectorScalingY = dpi / default_screen_dpi * val(GetITSetting("VectorScalingY", "1"))
    BitmapScalingX = val(GetITSetting("BitmapScalingX", "1"))
    BitmapScalingY = val(GetITSetting("BitmapScalingY", "1"))
    
    ' Write latex to a temp file
    WriteToFile TempPath, FilePrefix, TextWindow1.Text
    
    ' Run latex
    #If Mac Then
        Dim fs As New MacFileSystemObject
    #Else
        Dim fs As New FileSystemObject
    #End If
    FrameProcess.Visible = True
    
    Dim RetVal As Long, RetValConv As Long
    Dim FinalFilename As String
    Dim ErrorMessage As String
    Dim RunCommand As String
    
    If UseVector = True And VectorOutputType = "tex2img" Then
        ' Use TeX2img to generate an EMF file from LaTeX
        LabelProcess.Caption = "LaTeX to EMF..."
        FrameProcess.Repaint
        RunCommand = ShellEscape(tex2img_command) & " --latex " + latex_command _
                            & " --preview- " + FilePrefix + ".tex" & " " + FilePrefix + ".emf"
        RetVal& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
        If (RetVal& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".emf")) Then
            ErrorMessage = "TeX2img did not return in " & TimeOutTimeString & " seconds and may have hung." _
                    & vbNewLine & "You should have run TeX2img once outside IguanaTex to make sure its path are set correctly." _
                    & vbNewLine & "Please make sure your code compiles outside IguanaTex."
            ShowError ErrorMessage, RunCommand
            FrameProcess.Visible = False
            Exit Sub
        End If
        FinalFilename = FilePrefix & ".emf"
        OutputType = "EMF"
    Else
        If UseDVI = True Then
            ' Convert to DVI
            If latex_command = "xelatex" Then
                OutputType = "XDV"
                OutputExt = ".xdv"
            Else
                OutputType = "DVI"
                OutputExt = ".dvi"
            End If
            LabelProcess.Caption = "LaTeX to " & OutputType & "..."
            FrameProcess.Repaint
            If UseLatexmk = True Then
                RunCommand = ShellEscape(TeXExePath & latexmk_command & TeXExeExt) & " " & latexmk_dvi_options _
                                    & " -shell-escape -interaction=batchmode " + FilePrefix + ".tex"
            Else ' Run latex engine in DVI/XDV output mode
                RunCommand = ShellEscape(TeXExePath & latex_command & TeXExeExt) & " " & latex_dvi_options _
                                    & " -shell-escape -interaction=batchmode " & FilePrefix + ".tex"
            End If
            RetVal& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
            If (RetVal& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & OutputExt)) Then
                ' Error in Latex code
                ' Read log file and show it to the user
                If fs.FileExists(TempPath & FilePrefix & ".log") Then
                    ShowLogFile (TempPath + FilePrefix + ".log")
                Else
                    ErrorMessage = "latex did not return in " & TimeOutTimeString & " seconds and may have hung." _
                    & vbNewLine & "Please make sure your code compiles outside IguanaTex." _
                    & vbNewLine & "You may also try generating in Debug mode, as it will let you know if any font/package is missing."
                    ShowError ErrorMessage, RunCommand
                End If
                FrameProcess.Visible = False
                Exit Sub
            End If
            
            If UsePDF = True Then
                ' Further convert to PDF
                LabelProcess.Caption = OutputType & " to PDF..."
                FrameProcess.Repaint
                RunCommand = ShellEscape(TeXExePath & "dvipdfmx" & TeXExeExt) & " -o " + FilePrefix + ".pdf" _
                                        & " " & FilePrefix & OutputExt
                RetValConv& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
                If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".pdf")) Then
                    ' Error in DVI to PDF conversion
                    ErrorMessage = "Error while using dvipdfmx to convert from " & OutputType & " to PDF."
                    ShowError ErrorMessage, RunCommand
                    FrameProcess.Visible = False
                    Exit Sub
                End If
                OutputType = "PDF"
                OutputExt = ".pdf"
            End If
        Else ' If UseDVI is False, then UsePDF must be true: convert straight to PDF
            OutputType = "PDF"
            OutputExt = ".pdf"
            LabelProcess.Caption = "LaTeX to PDF..."
            FrameProcess.Repaint
            If UseLatexmk = True Then
                RunCommand = ShellEscape(TeXExePath & latexmk_command & TeXExeExt) & " " & latexmk_pdf_options _
                            & " -shell-escape -interaction=batchmode " & FilePrefix + ".tex"
            Else
                RunCommand = ShellEscape(TeXExePath & latex_command & TeXExeExt) & " -shell-escape -interaction=batchmode " _
                                        & FilePrefix + ".tex"
            End If
            RetVal& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
            
            If (RetVal& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & OutputExt)) Then
                ' Error in Latex code
                ' Read log file and show it to the user
                If fs.FileExists(TempPath & FilePrefix & ".log") Then
                    ShowLogFile (TempPath & FilePrefix & ".log")
                Else
                    ErrorMessage = latex_command & " did not return in " & TimeOutTimeString & " seconds and may have hung." _
                    & vbNewLine & "Please make sure your code compiles outside IguanaTex." _
                    & vbNewLine & "You may also try generating in Debug mode, as it will let you know if any font/package is missing"
                    ShowError ErrorMessage, RunCommand
                End If
                FrameProcess.Visible = False
                Exit Sub
            End If
        End If
        
        ' By now, we either have a DVI/XDV file, or a PDF file. Let's generate a Shape (EMF/SVG) or Picture (PNG/PDF) display.
        If UseVector Then
            ' Shape display -- formerly known as "Vector"
            ' I won't replace Vector with Shape everywhere in the code to avoid introducing bugs
            If VectorOutputType = "pdfiumdraw" Then
                ' Use pdfiumdraw to generate an EMF file from the previously generated PDF
                LabelProcess.Caption = "PDF to EMF..."
                FrameProcess.Repaint
                RunCommand = ShellEscape(pdfiumdraw_command) & " --extent=50 --emf --transparent --pages=1 " _
                                    & FilePrefix & ".pdf"
                RetVal& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
                If (RetVal& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".emf")) Then
                    ErrorMessage = "TeX2img's pdfiumdraw did not return in " & TimeOutTimeString & " seconds and may have hung." _
                            & vbNewLine & "You should have run TeX2img once outside IguanaTex to make sure its path are set correctly." _
                            & vbNewLine & "Please make sure your code compiles outside IguanaTex."
                    ShowError ErrorMessage, RunCommand
                    FrameProcess.Visible = False
                    Exit Sub
                End If
                FinalFilename = FilePrefix & ".emf"
                OutputType = "EMF"
            Else
                ' Use dvisvgm to generate SVG (either from DVI/XDV or from PDF)
                LabelProcess.Caption = OutputType & " to SVG..."
                FrameProcess.Repaint
                Dim dvisvgm_options As String
                If OutputType = "PDF" Then
                    dvisvgm_options = " --pdf"
                Else
                    dvisvgm_options = " --no-fonts"
                End If
                RunCommand = ShellEscape(TeXExePath & "dvisvgm" & TeXExeExt) & dvisvgm_options & " -o " _
                                    & FilePrefix & ".svg" & libgsString & " " _
                                    & FilePrefix & OutputExt
                RetValConv& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
                If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".svg")) Then
                    ' Error in DVI/XDV/PDF to SVG conversion
                    ErrorMessage = "Error while using dvisvgm to convert from " & OutputType & " to SVG."
                    ShowError ErrorMessage, RunCommand
                    FrameProcess.Visible = False
                    Exit Sub
                End If
                FinalFilename = FilePrefix & ".svg"
                OutputType = "SVG"
            
            End If
        Else
            ' Picture display: PDF or PNG on Mac, PNG on PC
            If PictureOutputType = "PDF" Then
                LabelProcess.Caption = "Cropping PDF..."
                FrameProcess.Repaint
            Else
                LabelProcess.Caption = OutputType & " to PNG..."
                FrameProcess.Repaint
            End If
            If OutputType = "PDF" Then ' Crop PDF and (on Windows) convert to PNG
                ' Output Bounding Box to file and read back in the appropriate information
                #If Mac Then
                    RunCommand = ShellEscape(gs_command) & " -q -dBATCH -dNOPAUSE -sDEVICE=bbox " _
                                        & FilePrefix & ".pdf" & " 2> " & FilePrefix & ".bbx"
                #Else
                    RunCommand = "cmd /C " & ShellEscape(gs_command) & " -q -dBATCH -dNOPAUSE -sDEVICE=bbox " _
                                            & FilePrefix & ".pdf" & " 2> " & FilePrefix & ".bbx"
                #End If
                RetValConv& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
                If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".bbx")) Then
                    ' Error in bounding box computation
                    ErrorMessage = "Error while using Ghostscript to compute the bounding box. Is your path correct?"
                    ShowError ErrorMessage, RunCommand
                    FrameProcess.Visible = False
                    Exit Sub
                End If
                Dim BBString As String
                BBString = BoundingBoxString(TempPath + FilePrefix + ".bbx")
                
                If PictureOutputType = "PDF" Then ' Only on Mac
                    ' PDF insert supported on Mac, only need to crop
                    RunCommand = ShellEscape(gs_command) & " -q -dBATCH -dNOPAUSE -sDEVICE=pdfwrite -sOutputFile=" _
                                        & FilePrefix & "_tmp.pdf" & BBString _
                                        & " -f " & FilePrefix & ".pdf"
                    RetValConv& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
                    If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & "_tmp.pdf")) Then
                        ' Error in PDF crop
                        ErrorMessage = "Error while using Ghostscript to crop the PDF. Is your path correct?"
                        ShowError ErrorMessage, RunCommand
                        FrameProcess.Visible = False
                        Exit Sub
                    End If
                    OutputType = "PDF"
                    FinalFilename = FilePrefix & "_tmp.pdf"
                Else ' This should only occur on Win, because we force DVI->PNG conversion on Mac for PNG
                    ' Convert PDF to PNG
                    RunCommand = ShellEscape(gs_command) & " -q -dBATCH -dNOPAUSE -sDEVICE=pngalpha -r" & OutputDpiString _
                                        & " -sOutputFile=" & FilePrefix & "_tmp.png" & BBString _
                                        & " -f " & FilePrefix & ".pdf"
                    RetValConv& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
                    If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & "_tmp.png")) Then
                        ' Error in PDF to PNG conversion
                        ErrorMessage = "Error while using Ghostscript to convert from PDF to PNG. Is your path correct?"
                        ShowError ErrorMessage, RunCommand
                        FrameProcess.Visible = False
                        Exit Sub
                    End If
                    ' Unfortunately, the resulting file has a metadata DPI of OutputDpi (=1200), not the default screen one (usually 96),
                    ' so there is a discrepancy with the dvipng output, which is always 96 (independent of the screen, actually).
                    ' The only workaround I have found so far is to use Imagemagick's convert to change the DPI (but not the pixel size!)
                    RunCommand = ShellEscape(IMconv) & " -units PixelsPerInch " & FilePrefix & "_tmp.png" _
                                            & " -density " & CStr(default_screen_dpi) & " " & FilePrefix & ".png"
                    RetValConv& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
                    If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".png")) Then
                        ' Error in PDF to PNG conversion
                        ErrorMessage = "Error while using ImageMagick to change the PNG DPI. Is your path correct?" _
                        & vbNewLine & "The full path is needed to avoid conflict with Windows's built-in convert.exe."
                        ShowError ErrorMessage, RunCommand
                        FrameProcess.Visible = False
                        Exit Sub
                    End If
                    ' 'I considered using ImageMagick's convert, but it's extremely slow, and uses ghostscript in the backend anyway
                    'PdfPngSwitches = "-density 1200 -trim -transparent white -antialias +repage"
                    'Execute IMconv & " " & PdfPngSwitches & " """ & FilePrefix & ".pdf"" """ & FilePrefix & ".png""", TempPath, debugMode
                    OutputType = "PNG"
                    FinalFilename = FilePrefix & ".png"
                End If
            Else ' Convert DVI to PNG
                Dim DviPngSwitches As String
                ' monitor is 96 dpi or higher; we use OutputDpi (=1200 by default) dpi to get a crisper display,
                ' and rescale later on for new displays to match the point size
                DviPngSwitches = "-q -D " & OutputDpiString & " -T tight -bg Transparent"
                ' If the user created a .png by using the standalone class with convert, we use that, else we use dvipng
                If Not fs.FileExists(TempPath & FilePrefix & ".png") Then
                    RunCommand = ShellEscape(TeXExePath & "dvipng" & TeXExeExt) & " " & DviPngSwitches _
                                        & " -o " & FilePrefix & ".png" & " " & FilePrefix & ".dvi"
                    RetValConv& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
                    If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".png")) Then
                        ErrorMessage = "dvipng failed, or did not return in " & TimeOutTimeString & " seconds and may have hung." _
                            & vbNewLine & "You may also try generating in Debug mode, as it will let you know if any font is missing."
                        ShowError ErrorMessage, RunCommand
                        FrameProcess.Visible = False
                        Exit Sub
                    End If
                End If
                OutputType = "PNG"
                FinalFilename = FilePrefix & ".png"
            End If
        End If
    End If
    
    ' Latex run successful.
    
    
    ' Now we prepare the insertion of the image
    LabelProcess.Caption = "Insert image..."
    FrameProcess.Repaint
    
    ' If we are in Edit mode, store parameters of old image
    Dim posX As Single
    Dim posY As Single
    Dim oldHeight As Single
    Dim oldWidth As Single
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim oldshape As Shape
    Dim oldshapeIsVector As Boolean
    Dim s As Shape
    Dim j As Long
    Dim IsInGroup As Boolean
    IsInGroup = False
    If ButtonRun.Caption = "ReGenerate" Then
        If Sel.ShapeRange.Type = msoGroup And Sel.HasChildShapeRange Then
            ' Old image is part of a group
            Set oldshape = Sel.ChildShapeRange(1)
            IsInGroup = True
            Dim arr() As Variant ' gather all shapes to be regrouped later on
            j = 0
            For Each s In Sel.ShapeRange.GroupItems
                If s.Name <> oldshape.Name Then
                    j = j + 1
                    ReDim Preserve arr(1 To j)
                    arr(j) = s.Name
                End If
            Next
            
            ' Store the group's animation and Zorder info in a dummy object tmpGroup
            Dim oldGroup As Shape
            Set oldGroup = Sel.ShapeRange(1)
            Dim tmpGroup As Shape
            Set tmpGroup = ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeDiamond, 1, 1, 1, 1)
            MoveAnimation oldGroup, tmpGroup
            MatchZOrder oldGroup, tmpGroup
            
            ' Tag all elements in the group with their hierarchy level and their name or group name
            Dim MaxGroupLevel As Long
            MaxGroupLevel = TagGroupHierarchy(arr, oldshape.Name)
            
        Else
            Set oldshape = Sel.ShapeRange(1)
        End If
        posX = oldshape.Left
        posY = oldshape.Top
        oldHeight = oldshape.Height
        oldWidth = oldshape.Width
        oldshapeIsVector = False
        If oldshape.Tags.item("BitmapVector") <> vbNullString Then
            If oldshape.Tags.item("BitmapVector") = 1 Then
                oldshapeIsVector = True
            End If
        End If
    Else
        posX = 200
        posY = 200
        If Sel.Type = ppSelectionShapes Then ' if something is selected on a slide, use its position for the new display
            'If Sel.ShapeRange.Type = msoGroup And Sel.HasChildShapeRange Then
            '    Set oldshape = Sel.ChildShapeRange(1)
            'Else
            '    Set oldshape = Sel.ShapeRange(1)
            'End If
            posX = Sel.ShapeRange(1).Left
            posY = Sel.ShapeRange(1).Top
        End If
    End If
            
    ' Get scaling factors
    Dim isTexpoint As Boolean
    Dim tScaleWidth As Single, tScaleHeight As Single
    Dim MagicScalingFactorEMF As Single, MagicScalingFactorPNG As Single
    Dim MagicScalingFactorSVG As Single, MagicScalingFactorPDF As Single
    Dim MagicScalingFactor As Single
    MagicScalingFactorEMF = 1 ' 1 / 100 ' Magical scaling factor for EMF.
    MagicScalingFactorPNG = default_screen_dpi / OutputDpi
    MagicScalingFactorPDF = 1
    MagicScalingFactorSVG = 1
    Select Case OutputType
        Case "EMF": MagicScalingFactor = MagicScalingFactorEMF
        Case "PNG": MagicScalingFactor = MagicScalingFactorPNG
        Case "PDF": MagicScalingFactor = MagicScalingFactorPDF
        Case "SVG": MagicScalingFactor = MagicScalingFactorSVG
        Case Else: MagicScalingFactor = 1
    End Select
    
    Dim PointSize As Single
    If ButtonRun.Caption <> "ReGenerate" Or CheckBoxReset.value Then
        PointSize = val(textboxSize.Text)
        tScaleWidth = PointSize / 10 * MagicScalingFactor  ' 1/10 is for the default LaTeX point size (10 pt)
        tScaleHeight = tScaleWidth
    Else
        ' Handle the case of Texpoint displays
        isTexpoint = False
        Dim OldDpi As Long
        OldDpi = OutputDpi
        With oldshape.Tags
            If .item("TEXPOINTSCALING") <> vbNullString Then
                isTexpoint = True
                tScaleWidth = val(.item("TEXPOINTSCALING")) * MagicScalingFactor
                tScaleHeight = tScaleWidth
            End If
            If .item("OUTPUTDPI") <> vbNullString Then
                OldDpi = val(.item("OUTPUTDPI"))
            End If
        End With
        If Not isTexpoint Then ' modifying a normal display, either PNG or EMF
            Dim HeightOld As Single, WidthOld As Single
            HeightOld = oldshape.Height
            WidthOld = oldshape.Width
            tScaleHeight = 1
            tScaleWidth = 1
            If oldshapeIsVector = False Then ' this deals with displays from very old versions of IguanaTex that lack proper size tags
                oldshape.ScaleHeight 1#, msoTrue
                oldshape.ScaleWidth 1#, msoTrue
                tScaleHeight = HeightOld / oldshape.Height * 960 / OutputDpi ' 0.8=960/1200 is there to preserve scaling of displays created with old versions of IguanaTex
                tScaleWidth = WidthOld / oldshape.Width * 960 / OutputDpi
            End If
            With oldshape.Tags
                If .item("ORIGINALHEIGHT") <> vbNullString Then
                    Dim tmpHeight As Single
                    tmpHeight = val(.item("ORIGINALHEIGHT"))
                    tScaleHeight = HeightOld / tmpHeight * OldDpi / OutputDpi
                End If
                If .item("ORIGINALWIDTH") <> vbNullString Then
                    Dim tmpWidth As Single
                    tmpWidth = val(.item("ORIGINALWIDTH"))
                    tScaleWidth = WidthOld / tmpWidth * OldDpi / OutputDpi
                End If
            End With
            
            Dim OldMagicScalingFactor As Single
            OldMagicScalingFactor = 1
            With oldshape.Tags
                    If .item("OUTPUTTYPE") <> vbNullString Then
                        Select Case .item("OUTPUTTYPE")
                            Case "EMF": OldMagicScalingFactor = MagicScalingFactorEMF
                            Case "PNG": OldMagicScalingFactor = MagicScalingFactorPNG
                            Case "PDF": OldMagicScalingFactor = MagicScalingFactorPDF
                            Case "SVG": OldMagicScalingFactor = MagicScalingFactorSVG
                            Case Else: OldMagicScalingFactor = 1
                        End Select
                    Else ' from an older version where we do not record OutputType
                        If oldshapeIsVector = False Then ' PNG
                            OldMagicScalingFactor = MagicScalingFactorPNG
                        Else ' EMF
                            OldMagicScalingFactor = MagicScalingFactorEMF
                        End If
                    End If
            End With
            ' Compensate for any change between formats
            tScaleHeight = tScaleHeight * MagicScalingFactor / OldMagicScalingFactor
            tScaleWidth = tScaleWidth * MagicScalingFactor / OldMagicScalingFactor
'            If UseVector = True And oldshapeIsVector = False Then
'                tScaleHeight = tScaleHeight * MagicScalingFactorEMF / MagicScalingFactorPNG
'                tScaleWidth = tScaleWidth * MagicScalingFactorEMF / MagicScalingFactorPNG
'            ElseIf UseVector = False And oldshapeIsVector = True Then
'                tScaleHeight = tScaleHeight / MagicScalingFactorEMF * MagicScalingFactorPNG
'                tScaleWidth = tScaleWidth / MagicScalingFactorEMF * MagicScalingFactorPNG
'            End If
        End If
    End If
    
    
    ' Insert image and rescale it
    Dim NewShape As Shape
    Set NewShape = AddDisplayShape(TempPath + FinalFilename, posX, posY)
    
    If UseVector Then
        If OutputType = "EMF" Then
            ' Clean up, optionally rescale the EMF picture, and convert it into PPT object
            Set NewShape = ConvertEMF(NewShape, VectorScalingX * tScaleWidth, VectorScalingY * tScaleHeight, posX, posY, "emf", True, True)
        Else 'SVG case
            ' Clean up and convert SVG into PPT object
            Set NewShape = convertSVG(NewShape, tScaleWidth, tScaleHeight, posX, posY)
        End If
        ' Tag shape and its components with their "original" sizes,
        ' which we get by dividing their current height/width by the scaling factors applied above
        NewShape.Tags.Add "ORIGINALHEIGHT", NewShape.Height / tScaleHeight
        NewShape.Tags.Add "ORIGINALWIDTH", NewShape.Width / tScaleWidth
        If NewShape.Type = msoGroup Then
            For Each s In NewShape.GroupItems
                s.Tags.Add "ORIGINALHEIGHT", s.Height / tScaleHeight
                s.Tags.Add "ORIGINALWIDTH", s.Width / tScaleWidth
            Next
        End If
    Else
        ' Resize to the true size of the png file and adjust using the manual scaling factors set in Main Settings
        With NewShape
            .ScaleHeight 1#, msoTrue
            .ScaleWidth 1#, msoTrue
            .LockAspectRatio = msoFalse
            .ScaleHeight BitmapScalingY, msoFalse
            .ScaleWidth BitmapScalingX, msoFalse
            .Tags.Add "OUTPUTDPI", OutputDpi ' Stores this display's resolution
            ' Add tags storing the original height and width, used next time to keep resizing ratio.
            .Tags.Add "ORIGINALHEIGHT", NewShape.Height
            .Tags.Add "ORIGINALWIDTH", NewShape.Width
            ' Apply scaling factors
            .ScaleHeight tScaleHeight, msoFalse
            .ScaleWidth tScaleWidth, msoFalse
            .LockAspectRatio = msoTrue
        End With
    End If
    ' in v1.59, we start tagging the type of output the shape was obtained from, and the IguanaTex version number
    NewShape.Tags.Add "OUTPUTTYPE", OutputType
    NewShape.Tags.Add "IGUANATEXVERSION", IGUANATEX_VERSION
    
    If CheckBoxForcePreserveSize.value Then
        ' We are forcing the new shape to have the same size as the old shape
        ' This is useful when converting between Bitmap and Vector
        With NewShape
            .LockAspectRatio = msoFalse
            .Height = oldHeight
            .Width = oldWidth
            .LockAspectRatio = msoTrue
        End With
    End If
        
    If ButtonRun.Caption = "ReGenerate" Then ' We are editing+resetting size of an old display, we keep rotation
        NewShape.Rotation = oldshape.Rotation
        If Not CheckBoxReset.value Then
            NewShape.LockAspectRatio = oldshape.LockAspectRatio ' Unlock aspect ratio if old display had it unlocked
        End If
    End If
    
    ' Add tags
    AddTagsToShape NewShape
    If UseVector = True And NewShape.Type = msoGroup Then
        Set s = NewShape.GroupItems(1)
        AddTagsToShape s 'only left most for now, to make things simple
        For Each s In NewShape.GroupItems
        '    Call AddTagsToShape(s)
            s.Tags.Add "EMFchild", True
        Next
    End If
    
    ' Copy animation settings and formatting from old image, then delete it
    If ButtonRun.Caption = "ReGenerate" Then
        Dim TransferDesign As Boolean
        TransferDesign = True
        If UseVector <> oldshapeIsVector Or CheckBoxResetFormat.value Then
            TransferDesign = False
        End If
        Dim j_remain As Long, j_current As Long
        Dim n As Variant
        Dim ThisShapeLevel As Long, i_tag As Long, Level As Long
        If IsInGroup Then
            ' Transfer format to new shape
            MatchZOrder oldshape, NewShape
            If TransferDesign Then
                oldshape.PickUp
                NewShape.Apply
            End If
            ' Handle the case of shape within EMF group.
            Dim DeleteLowestLayer As Boolean
            DeleteLowestLayer = False
            If oldshape.Tags.item("EMFchild") <> vbNullString Then
                DeleteLowestLayer = True
            End If
            oldshape.Delete
            
            Dim newGroup As Shape
            ' Get current slide, it will be used to group ranges
            Dim sld As Slide
            Dim SlideIndex As Long
            SlideIndex = ActiveWindow.View.Slide.SlideIndex
            Set sld = ActivePresentation.Slides(SlideIndex)

            ' Group all non-modified elements from old group, plus modified element
            j = j + 1
            ReDim Preserve arr(1 To j)
            arr(j) = NewShape.Name
            If DeleteLowestLayer Then
                Dim arr_remain() As Variant
                j_remain = 0
                For Each n In arr
                    Set s = ActiveWindow.Selection.SlideRange.Shapes(n)
                    ThisShapeLevel = 0
                    For i_tag = 1 To s.Tags.count
                        If (s.Tags.Name(i_tag) = "LAYER") Then
                            ThisShapeLevel = val(s.Tags.value(i_tag))
                        End If
                    Next
                    If ThisShapeLevel = 1 Then
                        s.Delete
                    Else
                        j_remain = j_remain + 1
                        ReDim Preserve arr_remain(1 To j_remain)
                        arr_remain(j_remain) = s.Name
                    End If
                Next
                NewShape.Tags.Add "LAYER", 2
                arr = arr_remain
            Else
                NewShape.Tags.Add "LAYER", 1
            End If
            NewShape.Tags.Add "SELECTIONNAME", NewShape.Name
            
            ' Hierarchically re-group elements
            For Level = 1 To MaxGroupLevel
                Dim CurrentLevelArr() As Variant
                j_current = 0
                For Each n In arr
                    ThisShapeLevel = 0
                    Dim ThisShapeSelectionName As String
                    ThisShapeSelectionName = vbNullString
                    On Error Resume Next
                    With ActiveWindow.Selection.SlideRange.Shapes(n).Tags
                        For i_tag = 1 To .count
                            If (.Name(i_tag) = "LAYER") Then
                                ThisShapeLevel = val(.value(i_tag))
                            End If
                            If (.Name(i_tag) = "SELECTIONNAME") Then
                                ThisShapeSelectionName = .value(i_tag)
                            End If
                        Next
                    End With
                    
                    
                    If ThisShapeLevel = Level Then
                        If j_current > 0 Then
                            If Not IsInArray(CurrentLevelArr, ThisShapeSelectionName) Then
                                j_current = j_current + 1
                                ReDim Preserve CurrentLevelArr(1 To j_current)
                                CurrentLevelArr(j_current) = ThisShapeSelectionName
                            End If
                        Else
                            j_current = j_current + 1
                            ReDim Preserve CurrentLevelArr(1 To j_current)
                            CurrentLevelArr(j_current) = ThisShapeSelectionName
                        End If
                    End If
                Next
                
                If j_current > 1 Then
                    Set newGroup = sld.Shapes.Range(CurrentLevelArr).Group
                    j = j + 1
                    ReDim Preserve arr(1 To j)
                    arr(j) = newGroup.Name
                    newGroup.Tags.Add "SELECTIONNAME", newGroup.Name
                    newGroup.Tags.Add "LAYER", Level + 1
                End If
                
            Next
            
            ' Delete the tags to avoid conflict with future runs
            For Each n In arr
                On Error Resume Next
                    ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Delete ("SELECTIONNAME")
                    ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Delete ("LAYER")
            Next
            
            ' Use temporary group to retrieve the group's original animation and Zorder
            MoveAnimation tmpGroup, newGroup
            MatchZOrder tmpGroup, newGroup
            tmpGroup.Delete
        Else
            MoveAnimation oldshape, NewShape
            MatchZOrder oldshape, NewShape
            If TransferDesign Then
                If oldshapeIsVector And oldshape.Type = msoGroup Then
                    
                    ' First transfer group format to temporary shape
                    ' (we use a duplicate of the old EMF shape)
                    Dim tmpGroupEMF As Shape
                    Set tmpGroupEMF = oldshape.Duplicate(1)
                    'TransferGroupFormat oldshape, tmpGroupEMF
                    
                    ' Transfer shape formatting
                    ' First need to delete all but one shape in the group to unlock the format pickup
                    Dim tmpshp As Shape
                    Set tmpshp = oldshape.GroupItems(1)
                    For j = oldshape.GroupItems.count To 2 Step -1
                        ' Delete backwards because Powerpoint renumbers
                        ' We could also always delete .GroupItems(2) ...
                        oldshape.GroupItems(j).Delete
                    Next
                    tmpshp.PickUp
                    
                    ' Transfer shape formatting to each shape within the group
                    For Each s In NewShape.GroupItems
                        s.Apply
                    Next
                    tmpshp.Delete
                    
                    ' Now we can transfer the group formatting from the temporary shape
                    TransferGroupFormat tmpGroupEMF, NewShape
                    tmpGroupEMF.Delete
                Else
                    oldshape.PickUp
                    NewShape.Apply
                    oldshape.Delete
                End If
            Else
                oldshape.Delete
            End If
        End If
    End If
    
    ' Handle non-transparent case if selected by user in form; this used to be handled by making the PNG itself non-transparent
    If Not checkboxTransp.value Then
        NewShape.Fill.ForeColor.RGB = vbWhite
        NewShape.Fill.Visible = True
    End If
    
    ' Add Alternative Text
    If AddAltText = True Then
        NewShape.AlternativeText = TextWindow1.Text
        If UseVector = True Then
            NewShape.Title = "IguanaTex Shape Display"
        Else
            NewShape.Title = "IguanaTex Picture Display"
        End If
    End If
    
    ' Select the new shape
    NewShape.Select
    
    
    ' Delete temp files if not in debug mode, external editor not used, and chose not to keep them in Main Settings
    Dim KeepTempFiles As Boolean
    KeepTempFiles = GetITSetting("KeepTempFiles", True)
    If (Not debugMode) And (Not UseExternalEditor) And (Not KeepTempFiles) Then
        #If Mac Then
            fs.FindDelete TempPath, FilePrefix + "*.*"
        #Else
            fs.DeleteFile TempPath + FilePrefix + "*.*"
        #End If
    End If
    
    
    
    FrameProcess.Visible = False
    Unload LatexForm
Exit Sub
   
End Sub

Private Sub AddTagsToShape(ByVal vSh As Shape)
    With vSh.Tags
        .Add "LATEXADDIN", TextWindow1.Text
        .Add "IguanaTexSize", val(textboxSize.Text)
        .Add "IGUANATEXCURSOR", TextWindow1.SelStart
        .Add "TRANSPARENCY", checkboxTransp.value
        .Add "FILENAME", TextBoxFile.Text
        .Add "LATEXENGINEID", ComboBoxLaTexEngine.ListIndex
        .Add "TEMPFOLDER", TextBoxTempFolder.Text
        .Add "LATEXFORMHEIGHT", LatexForm.Height
        .Add "LATEXFORMWIDTH", LatexForm.Width
        .Add "LATEXFORMWRAP", TextWindow1.WordWrap
        .Add "BitmapVector", ComboBoxBitmapVector.ListIndex
    End With
End Sub

Private Sub ShowLogFile(LogFileName As String)
    LogFileViewer.TextBox1.Text = ReadAll(LogFileName)
    LogFileViewer.TextBox1.ScrollBars = fmScrollBarsBoth
    LogFileViewer.Show 1
End Sub


Private Function IsInArray(ByVal arr As Variant, ByVal valueToCheck As String) As Boolean
    IsInArray = False
    Dim n As Variant
    For Each n In arr
        If n = valueToCheck Then
            IsInArray = True
            Exit For
        End If
    Next

End Function

Private Function TagGroupHierarchy(arr As Variant, TargetName As String) As Long
    ' Arr is the list of names of (leaf) elements in this group
    ' TargetName is the display which is being modified. We're going down the branch containing it.
    Dim Sel As Selection
    ActiveWindow.Selection.SlideRange.Shapes(TargetName).Select
    Set Sel = Application.ActiveWindow.Selection
    
    ' This function expects to receive a grouped ShapeRange
    ' We ungroup to reveal the structure at the layer below
    Sel.ShapeRange.Ungroup
    ActiveWindow.Selection.SlideRange.Shapes(TargetName).Select
           
    If Sel.ShapeRange.Type = msoGroup Then
        ' We need to go further down, the element being edited is still within a group
        ' Get the name of the Target group in which it is
        Dim TargetGroupName As String
        TargetGroupName = Sel.ShapeRange(1).Name
        
        Dim Arr_In() As Variant ' shapes in the same group
        Dim Arr_Out() As Variant ' shapes not in the same group
        
        ' Split range according to whether elements are in the same group or not
        Dim j_in As Long
        Dim j_out As Long
        Dim n As Variant
        j_in = 0
        j_out = 0
        For Each n In arr
            ActiveWindow.Selection.SlideRange.Shapes(n).Select
            If Sel.ShapeRange.Type = msoGroup Then
                ' object is in group
                If Sel.ShapeRange(1).Name = TargetGroupName Then
                    j_in = j_in + 1
                    ReDim Preserve Arr_In(1 To j_in)
                    Arr_In(j_in) = n
                Else
                    j_out = j_out + 1
                    ReDim Preserve Arr_Out(1 To j_out)
                    Arr_Out(j_out) = n
                End If
            Else ' object not in group, so it can't be in the same group as Target
                j_out = j_out + 1
                ReDim Preserve Arr_Out(1 To j_out)
                Arr_Out(j_out) = n
            End If
        Next
        
        ' Build shape range with all elements in that group, go one level down
        Dim Tmp As Long
        Tmp = TagGroupHierarchy(Arr_In, TargetName)
        TagGroupHierarchy = Tmp + 1
        
        ' For all elements not in that group, tag them
        For Each n In Arr_Out
            ActiveWindow.Selection.SlideRange.Shapes(n).Select
            ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "LAYER", TagGroupHierarchy
            If Sel.ShapeRange.Type = msoGroup Then
                ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "SELECTIONNAME", Sel.ShapeRange(1).Name
            Else
                ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "SELECTIONNAME", n
            End If
        Next
        
    Else ' we reached the final layer: the element being edited is by itself,
         ' all other elements will need to be handled either through their group
         ' name if in a group, or their name if not
        TagGroupHierarchy = 1
        For Each n In arr
            ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "LAYER", TagGroupHierarchy
            ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "SELECTIONNAME", n
        Next
    End If


End Function

Private Function BoundingBoxString(ByVal BBXFile As String) As String
    Dim TextSplit As Variant
    Dim tmptext As String
    Dim llx As Double, lly As Double, urx As Double, ury As Double
    Dim sx As String, sy As String, cx As String, cy As String
    Dim OutputDpiString As String
    OutputDpiString = TextBoxLocalDPI.Text
    Dim OutputDpi As Long
    #If Mac Then
        OutputDpi = 720
        Dim fnum As Integer
        fnum = FreeFile()
        Open BBXFile For Input As #fnum
        While Not EOF(fnum)
        Line Input #fnum, tmptext
        TextSplit = Split(tmptext, " ")
        If TextSplit(0) = "%%HiResBoundingBox:" Then
            ' Without the +/- 0.1, we noticed that the crop was too tight
            ' On the other hand, not using the Hires BB results in wide margins (but that's the default in pdfcrop)
            ' On Mac, +/-0.1 looks great, but it results on Windows in a display that appears cropped
            ' (within a box of the same size...). So I tried adding +/-1 to be extra safe, but this leads to other
            ' issues: the size is still different on Windows, and it also messes up the scaling when vectorizing on
            ' the Mac. So I decided to revert to 0.1 until I can find a real fix. Windows users will need to
            ' "regenerate" displays that appear crop.
            ' Another option would be to use PNG on Mac as well.
            llx = val(TextSplit(1)) - 0.1
            lly = val(TextSplit(2)) - 0.1
            urx = val(TextSplit(3)) + 0.1
            ury = val(TextSplit(4)) + 0.1
            'compute size and offset
            sx = CStr(RoundUp((urx - llx) / 72 * OutputDpi))
            sy = CStr(RoundUp((ury - lly) / 72 * OutputDpi))
            cx = Str(-llx)
            cy = Str(-lly)
        End If
        Wend
        Close #fnum
        BoundingBoxString = " -g" & sx & "x" & sy & " -dFIXEDMEDIA -c ""<</PageOffset [" & cx & " " & cy & "]>>setpagedevice"""
        '" -c ""<</Install {" & cx & " " & cy & " translate}>> setpagedevice"""
    #Else
        OutputDpi = val(OutputDpiString)
        Const ForReading As Long = 1
        Dim fs As New FileSystemObject
        Dim txtStream As TextStream
        Set txtStream = fs.OpenTextFile(BBXFile, ForReading, False)
        Do While Not txtStream.AtEndOfStream
        tmptext = txtStream.ReadLine
        TextSplit = Split(tmptext, " ")
        If TextSplit(0) = "%%HiResBoundingBox:" Then
            llx = val(TextSplit(1)) - 0.1
            lly = val(TextSplit(2)) - 0.1
            urx = val(TextSplit(3)) + 0.1
            ury = val(TextSplit(4)) + 0.1
            'compute size and offset
            sx = CStr(Round((urx - llx) / 72 * OutputDpi))
            sy = CStr(Round((ury - lly) / 72 * OutputDpi))
            cx = Str$(-llx)
            cy = Str$(-lly)
        End If
        Loop
        txtStream.Close
        BoundingBoxString = " -g" & sx & "x" & sy & " -c ""<</Install {" & cx & " " & cy & " translate}>> setpagedevice"""
    #End If
    

    
End Function

Private Sub SaveSettings()
    SetITSetting "Transparent", REG_DWORD, BoolToInt(checkboxTransp.value)
    SetITSetting "Debug", REG_DWORD, BoolToInt(checkboxDebug.value)
    SetITSetting "PointSize", REG_DWORD, CLng(val(textboxSize.Text))
    SetITSetting "LatexCode", REG_SZ, CStr(TextWindow1.Text)
    SetITSetting "LatexCodeCursor", REG_DWORD, CLng(TextWindow1.SelStart)
    #If Mac Then
        ' We save the height/width settings without the Mac resizing factor
        ' But until we make the window resizable, we don't save these settings,
        ' and instead let the user set them in Main Settings
        ' SetITSetting "LatexFormHeight", REG_DWORD, CLng(LatexForm.Height / gUserFormResizeFactor)
        ' SetITSetting "LatexFormWidth", REG_DWORD, CLng(LatexForm.Width / gUserFormResizeFactor)
    #Else
        SetITSetting "LatexFormHeight", REG_DWORD, CLng(LatexForm.Height)
        SetITSetting "LatexFormWidth", REG_DWORD, CLng(LatexForm.Width)
    #End If
    SetITSetting "EditorFontSize", REG_DWORD, CLng(TextWindow1.Font.Size)
    SetITSetting "Multipage", REG_SZ, MultiPage1.value
    SetITSetting "LatexFormWrap", REG_DWORD, BoolToInt(TextWindow1.WordWrap)
    'SetITSetting "EMFoutput", REG_DWORD, BoolToInt(CheckBoxEMF.Value)
    SetITSetting "BitmapVector", REG_DWORD, ComboBoxBitmapVector.ListIndex
    SetITSetting "OutputDpi", REG_DWORD, CLng(val(TextBoxLocalDPI.Text))
    
    
End Sub

Private Sub LoadSettings()
    checkboxTransp.value = CBool(GetITSetting("Transparent", True))
    checkboxDebug.value = CBool(GetITSetting("Debug", False))
    textboxSize.Text = GetITSetting("PointSize", "20")
    TextWindow1.Text = GetITSetting("LatexCode", DEFAULT_LATEX_CODE)
    TextWindow1.SelStart = GetITSetting("LatexCodeCursor", 0)
    MultiPage1.value = GetITSetting("Multipage", 0)
    TextWindow1.Font.Size = val(GetITSetting("EditorFontSize", "10"))
    TextBoxTempFolder.Text = GetTempPath()
    TextWindow1.WordWrap = CBool(GetITSetting("LatexFormWrap", True))
    ToggleButtonWrap.value = TextWindow1.WordWrap
    
    LaTexEngineList = GetLaTexEngineList()
    LaTexDVIOptionsList = GetLatexDVIOptionsList()
    LatexmkPDFOptionsList = GetLatexmkPDFOptionsList()
    LatexmkDVIOptionsList = GetLatexmkDVIOptionsList()
    UsePDFList = GetUsePDFList()
    UseDVIList = GetUseDVIList()
    ComboBoxLaTexEngine.List = GetLaTexEngineDisplayList()
    ComboBoxLaTexEngine.ListIndex = GetITSetting("LaTeXEngineID", 0)
    TextBoxLocalDPI.Text = GetITSetting("OutputDpi", "1200")
    ComboBoxBitmapVector.List = GetBitmapVectorList
    ComboBoxBitmapVector.ListIndex = GetITSetting("BitmapVector", 0)
            
    TemplateSortedListString = GetITSetting("TemplateSortedList", "0")
    TemplateSortedList = UnpackStringToArray(TemplateSortedListString)
    TemplateNameSortedListString = GetITSetting("TemplateNameSortedList", "New Template")
    ComboBoxTemplate.List = UnpackStringToArray(TemplateNameSortedListString)
End Sub


Sub ButtonTeXPath_Click()
    #If Mac Then
        TextBoxFile.Text = MacChooseFileOfType("tex")
    #Else
        TextBoxFile.Text = BrowseFilePath(TextBoxFile.Text, "Tex Files", "*.tex")
    #End If
    TextBoxFile.SetFocus
End Sub


Private Sub ComboBoxBitmapVector_Change()
    Apply_BitmapVector_Change
End Sub

Private Sub Apply_BitmapVector_Change()
    If ComboBoxBitmapVector.ListIndex = 1 Then
        checkboxTransp.Enabled = False
        checkboxTransp.value = True
        TextBoxLocalDPI.Enabled = False
        LabelDPI.Enabled = False
    Else
        checkboxTransp.Enabled = True
        TextBoxLocalDPI.Enabled = True
        LabelDPI.Enabled = True
    End If

End Sub

Sub CheckBoxReset_Click()
    Apply_CheckBoxReset
End Sub

Private Sub Apply_CheckBoxReset()
    textboxSize.Enabled = CheckBoxReset.value = True
End Sub

Sub CheckBoxForcePreserveSize_Click()
    If CheckBoxForcePreserveSize.value = True Then
        CheckBoxReset.Enabled = False
        CheckBoxReset.value = False
    Else
        CheckBoxReset.Enabled = True
    End If
    Apply_CheckBoxReset
End Sub

Sub ButtonAbout_Click()
    AboutBox.Show 1
End Sub


Sub ButtonMakeDefault_Click()
    SaveSettings
    Select Case MultiPage1.value
        Case 0 ' Direct input
            TextWindow1.SetFocus
        Case 1 ' Read from file
            TextBoxFile.SetFocus
        Case Else ' Templates
            TextWindowTemplateCode.SetFocus
    End Select
End Sub

Sub CmdButtonExternalEditor_Click()
    ExternalEditorForm.LaunchExternalEditor TextBoxTempFolder.Text, TextWindow1.Text
End Sub

Sub CmdButtonImportCode_Click()
    TextWindowTemplateCode.Text = TextWindow1.Text
    TextWindowTemplateCode.SelStart = TextWindow1.SelStart
    TextWindowTemplateCode.SetFocus
End Sub

Sub CmdButtonLoadTemplate_Click()
    If TextWindowTemplateCode.Text = vbNullString Then
        MsgBox "Please select a template to be loaded"
    Else
        TextWindow1.Text = TextWindowTemplateCode.Text
        TextWindow1.SelStart = TextWindowTemplateCode.SelStart
        MultiPage1.value = 0
        ToggleInputMode
    End If
End Sub

Sub CmdButtonRemoveTemplate_Click()
    Dim RemovedIndex As Long
    RemovedIndex = ComboBoxTemplate.ListIndex
    If ComboBoxTemplate.ListCount > 1 Then
        ' We should also be deleting the registry entry, but well, it does not take much space and will likely get reused anyway
        ComboBoxTemplate.RemoveItem RemovedIndex
        
        ' update the array that contains the sorted list of template IDs
        Dim templateID As Long
        templateID = TemplateSortedList(RemovedIndex)
        Dim i As Long
        For i = RemovedIndex To UBound(TemplateSortedList) - 1
            TemplateSortedList(i) = TemplateSortedList(i + 1)
        Next i
        TemplateSortedList(UBound(TemplateSortedList)) = templateID
        'NumberOfTemplates = NumberOfTemplates - 1
    Else
        ComboBoxTemplate.Clear
        ComboBoxTemplate.AddItem "New Template" 'prepare spot for new template
        ComboBoxTemplate.Text = vbNullString
        'NumberOfTemplates = 1
    End If
    UpdateTemplateRegistry
    ComboBoxTemplate.ListIndex = RemovedIndex
    TextWindowTemplateCode.SetFocus
End Sub

Sub CmdButtonSaveTemplate_Click()
    ' get the right ID from the array of sorted template IDs
    Dim templateID As Long
    templateID = TemplateSortedList(ComboBoxTemplate.ListIndex)
    ' add trailing new line if there isn't one: this helps with a bug where text with multi-byte characters gets chopped
    If Not Right$(TextWindowTemplateCode.Text, 1) = NEWLINE And Not Right$(TextWindowTemplateCode.Text, 1) = Chr$(10) Then
        TextWindowTemplateCode.Text = TextWindowTemplateCode.Text & NEWLINE
    End If
    ' build the corresponding registry key string
    ' Save name, code, and LaTeXEngineID
    Dim RegStr As String
    RegStr = "TemplateCode" & templateID
    SetITSetting RegStr, REG_SZ, CStr(TextWindowTemplateCode.Text)
    RegStr = "TemplateCodeSelStart" & templateID
    SetITSetting RegStr, REG_DWORD, CLng(TextWindowTemplateCode.SelStart)
    RegStr = "TemplateLaTeXEngineID" & templateID
    SetITSetting RegStr, REG_DWORD, ComboBoxLaTexEngine.ListIndex
    RegStr = "TemplateBitmapVector" & templateID
    SetITSetting RegStr, REG_DWORD, ComboBoxBitmapVector.ListIndex
    RegStr = "TemplateTempFolder" & templateID
    SetITSetting RegStr, REG_SZ, CStr(TextBoxTempFolder.Text)
    RegStr = "TemplateDPI" & templateID
    SetITSetting RegStr, REG_SZ, CStr(TextBoxLocalDPI.Text)
    ' if saved template was the "New Template", prepare new spot for next new template
    If ComboBoxTemplate.ListIndex = ComboBoxTemplate.ListCount - 1 Then
        ComboBoxTemplate.AddItem "New Template"
        'NumberOfTemplates = NumberOfTemplates + 1
        If ComboBoxTemplate.ListCount - 1 > UBound(TemplateSortedList) Then
            ReDim Preserve TemplateSortedList(0 To UBound(TemplateSortedList) + 1) As String
            TemplateSortedList(UBound(TemplateSortedList)) = CStr(ComboBoxTemplate.ListCount - 1)
        End If
    End If
    ComboBoxTemplate.List(ComboBoxTemplate.ListIndex) = TextBoxTemplateName.Text
    UpdateTemplateRegistry
    TextWindowTemplateCode.SetFocus
End Sub



Sub ComboBoxTemplate_Click()
    TextBoxTemplateName.Text = ComboBoxTemplate.Text
    ' Except for the empty "New Template" slot, get the code and LaTeXEngineID setting from registry
    If ComboBoxTemplate.ListIndex = ComboBoxTemplate.ListCount - 1 Then
        TextWindowTemplateCode.Text = vbNullString
        ComboBoxLaTexEngine.ListIndex = GetITSetting("LaTeXEngineID", 0)
        TextBoxTempFolder.Text = GetTempPath()
    Else
        ' get the right ID from the array of sorted template IDs
        Dim templateID As Long
        templateID = TemplateSortedList(ComboBoxTemplate.ListIndex)
        ' build the corresponding registry key string
        Dim RegStr As String
        RegStr = "TemplateCode" & templateID
        TextWindowTemplateCode.Text = GetITSetting(RegStr, vbNullString)
        RegStr = "TemplateCodeSelStart" & templateID
        TextWindowTemplateCode.SelStart = GetITSetting(RegStr, 0)
        RegStr = "TemplateLaTeXEngineID" & templateID
        ComboBoxLaTexEngine.ListIndex = GetITSetting(RegStr, GetITSetting("LaTeXEngineID", 0))
        RegStr = "TemplateBitmapVector" & templateID
        ComboBoxBitmapVector.ListIndex = GetITSetting(RegStr, GetITSetting("BitmapVector", False))
        RegStr = "TemplateTempFolder" & templateID
        TextBoxTempFolder.Text = GetITSetting(RegStr, GetTempPath())
        RegStr = "TemplateDPI" & templateID
        TextBoxLocalDPI.Text = GetITSetting(RegStr, vbNullString)
        Apply_BitmapVector_Change
    End If
    TextWindowTemplateCode.SetFocus
End Sub

Private Sub UpdateTemplateRegistry()
    ' update the list of saved templates names in the registry (will be used to initialize combo box content)
    TemplateSortedListString = PackArrayToString(TemplateSortedList)
    SetITSetting "TemplateSortedList", REG_SZ, CStr(TemplateSortedListString)
    ' save list of template names to registry
    Dim myArray() As String
    ReDim myArray(0 To ComboBoxTemplate.ListCount - 1) As String
    Dim i As Long
    For i = LBound(myArray) To UBound(myArray)
        myArray(i) = ComboBoxTemplate.List(i)
    Next i
    TemplateNameSortedListString = PackArrayToString(myArray)
    SetITSetting "TemplateNameSortedList", REG_SZ, CStr(TemplateNameSortedListString)
End Sub

Private Sub CmdButtonTemplateFontDown_Click()
    If TextWindowTemplateCode.Font.Size > 4 Then
        TextWindowTemplateCode.Font.Size = TextWindowTemplateCode.Font.Size - 1
    End If
End Sub

Private Sub CmdButtonTemplateFontUp_Click()
    If TextWindowTemplateCode.Font.Size < 72 Then
        TextWindowTemplateCode.Font.Size = TextWindowTemplateCode.Font.Size + 1
    End If
End Sub

Private Sub CmdButtonEditorFontDown_Click()
    If TextWindow1.Font.Size > 4 Then
        TextWindow1.Font.Size = TextWindow1.Font.Size - 1
    End If
End Sub

Private Sub CmdButtonEditorFontUp_Click()
    If TextWindow1.Font.Size < 72 Then
        TextWindow1.Font.Size = TextWindow1.Font.Size + 1
    End If
End Sub

Private Sub ToggleButtonWrap_Click()
    TextWindow1.WordWrap = ToggleButtonWrap.value = True
End Sub

Private Sub UserForm_Initialize()
    #If Mac Then
        
    #Else
        Set TextWindow1 = Me.TextBox1
        Set TextWindowTemplateCode = Me.TextBoxTemplateCode
    #End If

    LoadSettings
    
    ' With multiple monitors, the "CenterOwner" option to open the UserForm in the center of the parent window
    ' does not seem to work, at least in Office 2010.
    ' The following code to manually place the UserForm somehow makes the "CenterOwner" option work.
    ' Remark: if used with the Manual placement option, it would place the window to the left, under the ribbon.
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    Me.Height = 320
    Me.Width = 385
    #If Mac Then
        ResizeUserForm Me
    #End If
    
    LatexForm.textboxSize.Visible = True
    LatexForm.Label2.Visible = True
    LatexForm.Label3.Visible = True

    FrameProcess.Visible = False
    
    
End Sub

Private Function isFormModeless() As Boolean

    On Error GoTo EH

    Me.Show vbModeless
    isFormModeless = True

    Exit Function

EH:
    isFormModeless = False

End Function


Private Sub UserForm_Activate()
    DoneWithActivation = False

    ' We have to be careful of the case where the edit window gets activated in vbModeless mode
    If Not isFormModeless Then
        'Execute macro to enable resizeability
        MakeFormResizable Me
        
        #If Mac Then
            MacEnableCopyPaste Me
            MacEnableAccelerators Me
            checkboxDebug.Accelerator = "E"
        #End If
        
        If Not FormHeightWidthSet Then
            #If Mac Then
                LatexForm.Height = GetITSetting("LatexFormHeight", 320) * gUserFormResizeFactor
                LatexForm.Width = GetITSetting("LatexFormWidth", 385) * gUserFormResizeFactor
            #Else
                LatexForm.Height = GetITSetting("LatexFormHeight", 320)
                LatexForm.Width = GetITSetting("LatexFormWidth", 385)
            #End If
        End If
        'ResizeForm
        ToggleInputMode
        DoneWithActivation = True
    End If

End Sub

Sub RetrieveOldShapeInfo(ByVal oldshape As Shape, ByVal mainText As String)
    CheckBoxReset.Visible = True
    CheckBoxReset.value = False
    CheckBoxResetFormat.Visible = True
    CheckBoxResetFormat.value = False
    CheckBoxForcePreserveSize.Visible = True
    CheckBoxForcePreserveSize.value = False
    Label2.Caption = "Reset size:"
    ButtonRun.Caption = "ReGenerate"
    ButtonRun.Accelerator = "G"
    
    TextWindow1.Text = mainText
    Dim CursorPosition As Long
    CursorPosition = Len(TextWindow1.Text)
                
    Dim FormHeightSet As Boolean
    Dim FormWidthSet As Boolean
    FormHeightSet = False
    FormWidthSet = False
     
    With oldshape.Tags
        If .item("IGUANATEXSIZE") <> vbNullString Then
            textboxSize.Text = .item("IGUANATEXSIZE")
        End If
        If .item("OUTPUTDPI") <> vbNullString Then
            TextBoxLocalDPI.Text = .item("OUTPUTDPI")
        End If
        If .item("BitmapVector") <> vbNullString Then
            ComboBoxBitmapVector.ListIndex = .item("BitmapVector")
        End If
        If .item("TRANSPARENCY") <> vbNullString Then
            checkboxTransp.value = SanitizeBoolean(.item("TRANSPARENCY"), True)
        ElseIf .item("TRANSPARENT") <> vbNullString Then
            checkboxTransp.value = SanitizeBoolean(.item("TRANSPARENT"), True)
        End If
        If .item("IGUANATEXCURSOR") <> vbNullString Then
            CursorPosition = .item("IGUANATEXCURSOR")
        End If
        If .item("LATEXENGINEID") <> vbNullString Then
            ComboBoxLaTexEngine.ListIndex = .item("LATEXENGINEID")
        End If
        If .item("LATEXFORMHEIGHT") <> vbNullString Then
            LatexForm.Height = .item("LATEXFORMHEIGHT")
            FormHeightSet = True
        End If
        If .item("LATEXFORMWIDTH") <> vbNullString Then
            LatexForm.Width = .item("LATEXFORMWIDTH")
            FormWidthSet = True
        End If
        If .item("LATEXFORMWRAP") <> vbNullString Then
            TextWindow1.WordWrap = SanitizeBoolean(.item("LATEXFORMWRAP"), True)
            ToggleButtonWrap.value = TextWindow1.WordWrap
        End If
    End With
    Apply_BitmapVector_Change
    FormHeightWidthSet = FormHeightSet And FormWidthSet
    TextWindow1.SelStart = CursorPosition
    textboxSize.Enabled = False
End Sub


Private Function SanitizeBoolean(ByVal Str As String, ByVal Def As Boolean) As Boolean
    On Error GoTo ErrWrongBoolean:
    SanitizeBoolean = CBool(Str)
    Exit Function
ErrWrongBoolean:
    SanitizeBoolean = Def
    Resume Next
End Function

Private Sub UserForm_Resize()
    Dim minLatexFormHeight As Double
    Dim minLatexFormWidth As Double
    ' Minimal size
    Select Case MultiPage1.value
     Case 0 ' Direct input
        minLatexFormHeight = MultiPage1.Top + 18 + 50 + 6 * ButtonAbout.Height
     Case 1 ' Read from file
        minLatexFormHeight = MultiPage1.Top + 18 + 50 + 6 * ButtonAbout.Height
     Case Else ' Templates
        minLatexFormHeight = MultiPage1.Top + 18 + 50 + 9 * ButtonAbout.Height
    End Select
    minLatexFormWidth = ButtonCancel.Left + ButtonCancel.Width + ButtonAbout.Width + 24
     
    If LatexForm.Height < minLatexFormHeight Then
        LatexForm.Height = minLatexFormHeight
    End If
    If LatexForm.Width < minLatexFormWidth Then
        LatexForm.Width = minLatexFormWidth
    End If
    
    ResizeForm
End Sub

Private Sub ResizeForm()
    Dim bordersize As Integer
    bordersize = 6
    
    MultiPage1.Left = bordersize
    MultiPage1.Top = bordersize
    MultiPage1.Width = LatexForm.Width - bordersize * 2
    MultiPage1.Height = LatexForm.Height - MultiPage1.Top - ButtonAbout.Height * 4
    ' Here, we resize TextBox1 as our reference, not TextWindow1 (which is different on Mac, and gets resized to TextBox1)
    TextBox1.Width = MultiPage1.Width - 12
    TextBox1.Height = MultiPage1.Height - TextBox1.Top - 20
    
    'Other elements are moved as needed
    ButtonAbout.Top = MultiPage1.Top + MultiPage1.Height + bordersize
    ButtonAbout.Left = ButtonCancel.Left + ButtonCancel.Width + bordersize
    ButtonRun.Top = ButtonAbout.Top
    ButtonCancel.Top = ButtonRun.Top
    ButtonMakeDefault.Top = ButtonAbout.Top + ButtonAbout.Height + bordersize
    ButtonMakeDefault.Left = ButtonAbout.Left
    FrameProcess.Top = ButtonMakeDefault.Top
    FrameProcess.Left = ButtonRun.Left
    LabelProcess.Width = FrameProcess.Width
    LabelProcess.Top = 4
    
    CmdButtonImportCode.Left = TextBox1.Width - CmdButtonImportCode.Width
    CmdButtonLoadTemplate.Left = CmdButtonImportCode.Left
    CmdButtonSaveTemplate.Left = CmdButtonImportCode.Left
    CmdButtonRemoveTemplate.Left = CmdButtonImportCode.Left
    TextBoxTemplateCode.Width = CmdButtonImportCode.Left - bordersize
    TextBoxTemplateCode.Height = MultiPage1.Height - TextBoxTemplateCode.Top - 20
    
    textboxSize.Top = ButtonAbout.Top
    Label2.Top = textboxSize.Top + Round(textboxSize.Height - Label2.Height) / 2
    CheckBoxReset.Top = textboxSize.Top
    Label3.Top = Label2.Top
    checkboxTransp.Top = CheckBoxReset.Top + 21 'checkboxTransp.Height + 2
    CheckBoxForcePreserveSize.Top = checkboxTransp.Top
    CheckBoxForcePreserveSize.Left = checkboxTransp.Left + checkboxTransp.Width
    checkboxDebug.Top = checkboxTransp.Top + 21 ' checkboxTransp.Height + 2
    CheckBoxResetFormat.Top = checkboxDebug.Top
    CheckBoxResetFormat.Left = checkboxDebug.Left + checkboxDebug.Width + 10
    
    #If Mac Then
        TextWindow1.ResizeAsTarget
        TextWindowTemplateCode.ResizeAsTarget
    #End If

End Sub

Private Sub MoveAnimation(ByVal oldshape As Shape, ByVal NewShape As Shape)
    ' Move the animation settings of oldShape to newShape
    With ActiveWindow.Selection.SlideRange(1).TimeLine
        Dim eff As Effect
        For Each eff In .MainSequence
            If eff.Shape.Name = oldshape.Name Then eff.Shape = NewShape
        Next
    End With
End Sub

Private Sub MatchZOrder(ByVal oldshape As Shape, ByVal NewShape As Shape)
    ' Make the Z order of newShape equal to 1 higher than that of oldShape
    NewShape.ZOrder msoBringToFront
    While (NewShape.ZOrderPosition > oldshape.ZOrderPosition + 1)
        NewShape.ZOrder msoSendBackward
    Wend
End Sub

Private Sub DeleteAnimation(ByVal oldshape As Shape)
    ' Delete the animation settings of oldShape
    Dim i As Long
    With ActiveWindow.Selection.SlideRange(1).TimeLine
        For i = .MainSequence.count To 1 Step -1
            Dim eff As Effect
            Set eff = .MainSequence(i)
            If eff.Shape.Name = oldshape.Name Then eff.Delete
        Next
    End With
End Sub

Private Sub TransferGroupFormat(ByVal oldshape As Shape, ByRef NewShape As Shape)
    On Error Resume Next
    ' Transfer group formatting
    If oldshape.Glow.Radius > 0 Then
        NewShape.Glow.Color = oldshape.Glow.Color
        NewShape.Glow.Radius = oldshape.Glow.Radius
        NewShape.Glow.Transparency = oldshape.Glow.Transparency
    End If
    If oldshape.Reflection.Type <> msoReflectionTypeNone Then
        NewShape.Reflection.Blur = oldshape.Reflection.Blur
        NewShape.Reflection.Offset = oldshape.Reflection.Offset
        NewShape.Reflection.Size = oldshape.Reflection.Size
        NewShape.Reflection.Transparency = oldshape.Reflection.Transparency
        NewShape.Reflection.Type = oldshape.Reflection.Type
    End If
    
    If oldshape.SoftEdge.Type <> msoSoftEdgeTypeNone Then
        NewShape.SoftEdge.Radius = oldshape.SoftEdge.Radius
    End If
    
    If oldshape.Shadow.Visible Then
        NewShape.Shadow.Visible = oldshape.Shadow.Visible
        NewShape.Shadow.Blur = oldshape.Shadow.Blur
        NewShape.Shadow.ForeColor = oldshape.Shadow.ForeColor
        NewShape.Shadow.OffsetX = oldshape.Shadow.OffsetX
        NewShape.Shadow.OffsetY = oldshape.Shadow.OffsetY
        NewShape.Shadow.RotateWithShape = oldshape.Shadow.RotateWithShape
        NewShape.Shadow.Size = oldshape.Shadow.Size
        NewShape.Shadow.Style = oldshape.Shadow.Style
        NewShape.Shadow.Transparency = oldshape.Shadow.Transparency
        NewShape.Shadow.Type = oldshape.Shadow.Type
    End If
    
    If oldshape.ThreeD.Visible Then
        'newShape.ThreeD.BevelBottomDepth = oldshape.ThreeD.BevelBottomDepth
        'newShape.ThreeD.BevelBottomInset = oldshape.ThreeD.BevelBottomInset
        'newShape.ThreeD.BevelBottomType = oldshape.ThreeD.BevelBottomType
        'newShape.ThreeD.BevelTopDepth = oldshape.ThreeD.BevelTopDepth
        'newShape.ThreeD.BevelTopInset = oldshape.ThreeD.BevelTopInset
        'newShape.ThreeD.BevelTopType = oldshape.ThreeD.BevelTopType
        'newShape.ThreeD.ContourColor = oldshape.ThreeD.ContourColor
        'newShape.ThreeD.ContourWidth = oldshape.ThreeD.ContourWidth
        'newShape.ThreeD.Depth = oldshape.ThreeD.Depth
        'newShape.ThreeD.ExtrusionColor = oldshape.ThreeD.ExtrusionColor
        'newShape.ThreeD.ExtrusionColorType = oldshape.ThreeD.ExtrusionColorType
        NewShape.ThreeD.Visible = oldshape.ThreeD.Visible
        NewShape.ThreeD.Perspective = oldshape.ThreeD.Perspective
        NewShape.ThreeD.FieldOfView = oldshape.ThreeD.FieldOfView
        NewShape.ThreeD.LightAngle = oldshape.ThreeD.LightAngle
        'newShape.ThreeD.ProjectText = oldshape.ThreeD.ProjectText
        'If oldshape.ThreeD.PresetExtrusionDirection <> msoPresetExtrusionDirectionMixed Then
        '    newShape.ThreeD.SetExtrusionDirection oldshape.ThreeD.PresetExtrusionDirection
        'End If
        NewShape.ThreeD.PresetLighting = oldshape.ThreeD.PresetLighting
        If oldshape.ThreeD.PresetLightingDirection <> msoPresetLightingDirectionMixed Then
            NewShape.ThreeD.PresetLightingDirection = oldshape.ThreeD.PresetLightingDirection
        End If
        If oldshape.ThreeD.PresetLightingSoftness <> msoPresetLightingSoftnessMixed Then
            NewShape.ThreeD.PresetLightingSoftness = oldshape.ThreeD.PresetLightingSoftness
        End If
        If oldshape.ThreeD.PresetMaterial <> msoPresetMaterialMixed Then
            NewShape.ThreeD.PresetMaterial = oldshape.ThreeD.PresetMaterial
        End If
        If oldshape.ThreeD.PresetCamera <> msoPresetCameraMixed Then
            NewShape.ThreeD.SetPresetCamera oldshape.ThreeD.PresetCamera
        End If
        NewShape.ThreeD.RotationX = oldshape.ThreeD.RotationX
        NewShape.ThreeD.RotationY = oldshape.ThreeD.RotationY
        NewShape.ThreeD.RotationZ = oldshape.ThreeD.RotationZ
        'newShape.ThreeD.Z = oldshape.ThreeD.Z
    End If
End Sub

Private Function isTex(file As String) As Boolean
    isTex = GetExtension(file) = "tex"
End Function


Private Sub MultiPage1_Change()
    ToggleInputMode
End Sub


Private Sub TextBoxFile_Change()
    ButtonLoadFile.Enabled = FileExists(TextBoxFile.Text) And isTex(TextBoxFile.Text)
End Sub

Sub ButtonLoadFile_Click()
    MultiPage1.value = 0
    TextWindow1.Text = ReadAll(TextBoxFile.Text)
    ToggleInputMode
End Sub

Private Sub ToggleInputMode()
    #If Mac Then
        If MultiPage1.value = 0 Then
            TextWindow1.Show
            TextWindow1.SetResizeTarget TextBox1, Me
        Else
            TextWindow1.Hide
        End If
    
        If MultiPage1.value = 2 Then
            TextWindowTemplateCode.Show
            TextWindowTemplateCode.SetResizeTarget TextBoxTemplateCode, Me
        Else
            TextWindowTemplateCode.Hide
        End If
    #End If

    UserForm_Resize
    
    Select Case MultiPage1.value
        Case 0 ' Direct input
            TextWindow1.SetFocus
        Case 1 ' Read from file
            TextBoxFile.SetFocus
            ButtonLoadFile.Enabled = FileExists(TextBoxFile.Text) And isTex(TextBoxFile.Text)
        Case Else ' Templates
            If TextBoxTemplateName.Text = vbNullString Then
                TextBoxTemplateName.Text = ComboBoxTemplate.Text
            End If
            TextWindowTemplateCode.SetFocus
            
    End Select

    
End Sub

Private Function PackArrayToString(vArray As Variant) As String
    Dim strDelimiter As String
    strDelimiter = "|"
    PackArrayToString = Join(vArray, strDelimiter)
End Function

Private Function UnpackStringToArray(Str As String) As Variant
    Dim strDelimiter As String
    strDelimiter = "|"
    UnpackStringToArray = Split(Str, strDelimiter, , vbTextCompare)
End Function


' Attempt at getting DPI of previous display, but I cannot find a way to retrieve
' that info for an embedded display, as there does not seem to be a way to load
' the display as an ImageFile Object
' Requires Microsoft Windows Image Acquisition Library
'Private Function GetImageFileDPI(fileNm As String) As Long
'    Dim imgFile As Object
'    Set imgFile = CreateObject("WIA.ImageFile")
'    imgFile.LoadFile (fileNm)
'    GetImageFileDPI = 96
'    If imgFile.HorizontalResolution <> vbNullString Then
'        GetImageFileDPI = Round(imgFile.HorizontalResolution)
'    End If
'End Function

#If Mac Then

#Else
' Mousewheel functions

Private Sub TextBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal x As Single, ByVal Y As Single)
    If Not Me Is Nothing Then
        HookListBoxScroll Me, Me.TextBox1
    End If
End Sub

Private Sub TextBoxTemplateCode_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal x As Single, ByVal Y As Single)
    If Not Me Is Nothing Then
        HookListBoxScroll Me, Me.TextBoxTemplateCode
    End If
End Sub



' v1.58: I'm removing this because the support is not great, and I don't think scrolling is very useful
' for this combobox. The issue is that the combobox is within a frame, and once it gets the hook, we cannot
' unhook until we leave the whole frame, not just the combobox.
'Private Sub ComboBoxLaTexEngine_MouseMove( _
'                        ByVal Button As Integer, ByVal Shift As Integer, _
'                        ByVal X As Single, ByVal Y As Single)
'    If Not Me Is Nothing Then
'         HookListBoxScroll Me, Me.ComboBoxLaTexEngine
'    End If
'End Sub

' It seems difficult to get good mouse whell support simultaneously for the Bitmap/Vector combobox
' and the LatexEngine combobox, because they are in the same frame, and whoever gets the hook first holds to it.
'Private Sub ComboBoxBitmapVector_MouseMove( _
'                        ByVal Button As Integer, ByVal Shift As Integer, _
'                        ByVal X As Single, ByVal Y As Single)
'    If Not Me Is Nothing Then
'         HookListBoxScroll Me, Me.ComboBoxBitmapVector
'    End If
'End Sub


Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
        UnhookListBoxScroll
End Sub

#End If
