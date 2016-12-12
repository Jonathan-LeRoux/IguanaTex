VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LatexForm 
   Caption         =   "IguanaTex"
   ClientHeight    =   5880
   ClientLeft      =   14
   ClientTop       =   329
   ClientWidth     =   7560
   OleObjectBlob   =   "LatexForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LatexForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RegPath As String
    
Dim LaTexEngineList As Variant
Dim LaTexEngineDisplayList As Variant
Dim UsePDFList As Variant

'Dim NumberOfTemplates As Long
Dim TemplateSortedListString As String
Dim TemplateSortedList() As String
Dim TemplateNameSortedListString As String

Dim FormHeightWidthSet As Boolean

Dim theAppEventHandler As New AppEventHandler

Sub InitializeApp()
    Set theAppEventHandler.App = Application
    
    AddMenuItem "New Latex display...", "NewLatexEquation", 18 '226
    AddMenuItem "Edit Latex display...", "EditLatexEquation", 37
    AddMenuItem "Regenerate selected displays...", "RegenerateSelectedDisplaysNoChange", 19
    AddMenuItem "Convert to EMF...", "ConvertToEMF", 153
    AddMenuItem "Convert to PNG...", "ConvertToPNG", 931
    AddMenuItem "Settings...", "LoadSetTempForm", 548
    
End Sub

Sub AddMenuItem(itemText As String, itemCommand As String, itemFaceId As Long)
    ' Check if we have already added the menu item
    Dim initialized As Boolean
    Dim bef As Integer
    initialized = False
    bef = 1
    Dim Menu As CommandBars
    Set Menu = Application.CommandBars
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
        NewControl.Style = msoButton
    End If
End Sub

Sub UnInitializeApp()
    
    RemoveMenuItem "New Latex display..."
    RemoveMenuItem "Edit Latex display..."
    RemoveMenuItem "Regenerate selection..."
    RemoveMenuItem "Convert to EMF..."
    RemoveMenuItem "Convert to PNG..."
    RemoveMenuItem "Settings..."

End Sub

Sub RemoveMenuItem(itemText As String)
    Dim Menu As CommandBars
    Set Menu = Application.CommandBars
    For i = 1 To Menu("Insert").Controls.count
        If Menu("Insert").Controls(i).Caption = itemText Then
            Menu("Insert").Controls(i).Delete
            Exit For
        End If
    Next
    

End Sub


Private Sub ButtonCancel_Click()
    Unload LatexForm
    ' LatexForm.Hide
End Sub


Private Function IsPathWritable(TempPath As String) As Boolean
    FilePrefix = GetFilePrefix()
    
    Dim FName As String
    Dim FHdl As Integer
    FName = TempPath & FilePrefix & ".tmp"
    On Error GoTo TempFolderNotWritable
    FHdl = FreeFile()
    Open FName For Output Access Write As FHdl
    Print #FHdl, "TESTWRITE"
    Close FHdl
    IsPathWritable = True
    Kill FName
    
    On Error GoTo 0
    
    Exit Function

TempFolderNotWritable:
    IsPathWritable = False
End Function

Private Sub WriteLaTeX2File(TempPath As String, FilePrefix As String)
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(TempPath & FilePrefix & ".png") Then
        fs.DeleteFile TempPath + FilePrefix + "*.*" 'Make sure we don't keep old files
    End If
    RegPath = "Software\IguanaTex"
    Dim UseUTF8 As Boolean
    UseUTF8 = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "UseUTF8", True)
    
    If UseUTF8 = False Then
        Set f = fs.CreateTextFile(TempPath + FilePrefix + ".tex", True)
        f.Write TextBox1.Text
        f.Close
    Else
        Dim BinaryStream As Object
        Set BinaryStream = CreateObject("ADODB.stream")
        BinaryStream.Type = 1
        BinaryStream.Open
        Dim adodbStream  As Object
        Set adodbStream = CreateObject("ADODB.Stream")
        With adodbStream
            .Type = 2 'Stream type
            .Charset = "utf-8"
            .Open
            .WriteText TextBox1.Text
            '.SaveToFile TempPath & FilePrefix & ".tex", 2 'Save binary data To disk; problem: this includes a BOM
            ' Workaround to avoid BOM in file:
            .Position = 3 'skip BOM
            .CopyTo BinaryStream
            .Flush
            .Close
        End With
        BinaryStream.SaveToFile TempPath & FilePrefix & ".tex", 2 'Save binary data To disk
        BinaryStream.Flush
        BinaryStream.Close
    End If
    Set fs = Nothing
End Sub

Sub ButtonRun_Click()
    Dim TempPath As String
    'TempPath = GetTempPath()
    If Right(TextBoxTempFolder.Text, 1) <> "\" Then
        TextBoxTempFolder.Text = TextBoxTempFolder.Text & "\"
    End If
    TempPath = TextBoxTempFolder.Text
    
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
    
    Dim FilePrefix As String
    FilePrefix = GetFilePrefix()
    
    Dim debugMode As Boolean
    debugMode = checkboxDebug.Value
    
    ' Read settings
    RegPath = "Software\IguanaTex"
    LATEXENGINEID = ComboBoxLaTexEngine.ListIndex
    tex2pdf_command = LaTexEngineList(LATEXENGINEID)
    Dim UsePDF As Boolean
    UsePDF = UsePDFList(LATEXENGINEID)
    
    Dim UseEMF As Boolean
    BitmapVector = ComboBoxBitmapVector.ListIndex
    If BitmapVector = 0 Then
        UseEMF = False
    Else
        UseEMF = True
    End If
    'UseEMF = CheckBoxEMF.Value
    Dim OutputType As String
    
    Dim TimeOutTimeString As String
    Dim TimeOutTime As Long
    TimeOutTimeString = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TimeOutTime", "20") ' Wait 20 seconds for the processes to complete
    TimeOutTime = val(TimeOutTimeString) * 1000
    
    Dim OutputDpiString As String
    OutputDpiString = TextBoxLocalDPI.Text
    Dim OutputDpi As Long
    OutputDpi = val(OutputDpiString)
    
    ' Read current dpi in: this will be used when rescaling and optionally in pdf->png conversion
    dpi = lDotsPerInch
    default_screen_dpi = 96
    Dim VectorScalingX As Single, VectorScalingY As Single, BitmapScalingX As Single, BitmapScalingY As Single
    VectorScalingX = dpi / default_screen_dpi * val(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "VectorScalingX", "1"))
    VectorScalingY = dpi / default_screen_dpi * val(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "VectorScalingY", "1"))
    BitmapScalingX = val(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "BitmapScalingX", "1"))
    BitmapScalingY = val(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "BitmapScalingY", "1"))
        
    ' Test if path writable
    If Not IsPathWritable(TempPath) Then
        MsgBox "The temporary folder " & TempPath & " appears not to be writable."
        Exit Sub
    End If
    
    
    ' Write latex to a temp file
    Call WriteLaTeX2File(TempPath, FilePrefix)
    
    
    ' Run latex
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim LogFile As Object
    FrameProcess.Visible = True
    
    If UseEMF = True Then ' Use TeX2img to generate an EMF file
        tex2img_command = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TeX2img Command", "%USERPROFILE%\Downloads\TeX2img\TeX2imgc.exe")
        LabelProcess.Caption = "LaTeX to EMF..."
        FrameProcess.Repaint
        RetVal& = Execute("""" & tex2img_command & """ --latex " + tex2pdf_command + " --preview- """ + FilePrefix + ".tex"" """ + FilePrefix + ".emf""", TempPath, debugMode, TimeOutTime)
        If (RetVal& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".emf")) Then
            ' Error in Latex code
            ' Read log file and show it to the user
            If fs.FileExists(TempPath & FilePrefix & ".log") Then
                Set LogFile = fs.OpenTextFile(TempPath + FilePrefix + ".log", ForReading)
                LogFileViewer.TextBox1.Text = LogFile.ReadAll
                LogFile.Close
                LogFileViewer.TextBox1.ScrollBars = fmScrollBarsBoth
                LogFileViewer.Show 1
            Else
                MsgBox "TeX2img did not return in " & TimeOutTimeString & " seconds and may have hung." _
                & vbNewLine & "You should have run TeX2img once outside IguanaTex to make sure its path are set correctly." _
                & vbNewLine & "Please make sure your code compiles outside IguanaTex."
            End If
            FrameProcess.Visible = False
            Exit Sub
        End If
        FinalFilename = FilePrefix & ".emf"
        OutputType = "EMF"
    Else
        If UsePDF = True Then ' pdf to png route
            gs_command = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "GS Command", "C:\Program Files (x86)\gs\gs9.15\bin\gswin32c.exe")
            IMconv = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "IMconv", "C:\Program Files\ImageMagick\convert.exe")
            
            If tex2pdf_command = "platex" Then
                OutputExt = ".dvi"
                LabelProcess.Caption = "LaTeX to DVI..."
            Else
                OutputExt = ".pdf"
                LabelProcess.Caption = "LaTeX to PDF..."
            End If
            FrameProcess.Repaint
            
            RetVal& = Execute("""" & tex2pdf_command & """ -shell-escape -interaction=batchmode """ + FilePrefix + ".tex""", TempPath, debugMode, TimeOutTime)
            
            If (RetVal& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & OutputExt)) Then
                ' Error in Latex code
                ' Read log file and show it to the user
                If fs.FileExists(TempPath & FilePrefix & ".log") Then
                    Set LogFile = fs.OpenTextFile(TempPath + FilePrefix + ".log", ForReading)
                    LogFileViewer.TextBox1.Text = LogFile.ReadAll
                    LogFile.Close
                    LogFileViewer.TextBox1.ScrollBars = fmScrollBarsBoth
                    LogFileViewer.Show 1
                Else
                    MsgBox tex2pdf_command & " did not return in " & TimeOutTimeString & " seconds and may have hung." _
                    & vbNewLine & "Please make sure your code compiles outside IguanaTex." _
                    & vbNewLine & "You may also try generating in Debug mode, as it will let you know if any font/package is missing"
                End If
                FrameProcess.Visible = False
                Exit Sub
            End If
            
            If tex2pdf_command = "platex" Then
                LabelProcess.Caption = "DVI to PDF..."
                FrameProcess.Repaint
                ' platex actually outputs a DVI file, which we need to convert to PDF (we could go the EPS route, but this blends easier with IguanaTex's existing code)
                RetValConv& = Execute("dvipdfmx -o """ + FilePrefix + ".pdf"" """ & FilePrefix & ".dvi""", TempPath, debugMode, TimeOutTime)
                If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".pdf")) Then
                    ' Error in DVI to PDF conversion
                    MsgBox "Error while using dvipdm to convert from DVI to PDF."
                    FrameProcess.Visible = False
                    Exit Sub
                End If
            End If
            
            LabelProcess.Caption = "PDF to PNG..."
            FrameProcess.Repaint
            ' Output Bounding Box to file and read back in the appropriate information
            RetValConv& = Execute("cmd /C """ & gs_command & """ -q -dBATCH -dNOPAUSE -sDEVICE=bbox " & FilePrefix & ".pdf 2> " & FilePrefix & ".bbx", TempPath, debugMode, TimeOutTime)
            If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".bbx")) Then
                ' Error in bounding box computation
                MsgBox "Error while using Ghostscript to compute the bounding box. Is your path correct?"
                FrameProcess.Visible = False
                Exit Sub
            End If
            Dim BBString As String
            BBString = BoundingBoxString(TempPath + FilePrefix + ".bbx")
            
            ' Convert PDF to PNG
            If checkboxTransp.Value = True Then
                PdfPngDevice = "-sDEVICE=pngalpha"
            Else
                PdfPngDevice = "-sDEVICE=png16m"
            End If
            RetValConv& = Execute("""" & gs_command & """ -q -dBATCH -dNOPAUSE " & PdfPngDevice & " -r" & OutputDpiString & " -sOutputFile=""" & FilePrefix & "_tmp.png""" & BBString & " -f """ & TempPath & FilePrefix & ".pdf""", TempPath, debugMode, TimeOutTime)
            If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & "_tmp.png")) Then
                ' Error in PDF to PNG conversion
                MsgBox "Error while using Ghostscript to convert from PDF to PNG. Is your path correct?"
                FrameProcess.Visible = False
                Exit Sub
            End If
            ' Unfortunately, the resulting file has a metadata DPI of OutputDpi (=1200), not the default screen one (usually 96),
            ' so there is a discrepancy with the dvipng output, which is always 96 (independent of the screen, actually).
            ' The only workaround I have found so far is to use Imagemagick's convert to change the DPI (but not the pixel size!)
            ' Execute """" & IMconv & """ -units PixelsPerInch """ & FilePrefix & "_tmp.png"" -density " & CStr(dpi) & " """ & FilePrefix & ".png""", TempPath, debugMode
            RetValConv& = Execute("""" & IMconv & """ -units PixelsPerInch """ & FilePrefix & "_tmp.png"" -density " & CStr(default_screen_dpi) & " """ & FilePrefix & ".png""", TempPath, debugMode, TimeOutTime)
            If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".png")) Then
                ' Error in PDF to PNG conversion
                MsgBox "Error while using ImageMagick to change the PNG DPI. Is your path correct?" _
                & vbNewLine & "The full path is needed to avoid conflict with Windows's built-in convert.exe."
                FrameProcess.Visible = False
                Exit Sub
            End If
            
            ' 'I considered using ImageMagick's convert, but it's extremely slow, and uses ghostscript in the backend anyway
            'PdfPngSwitches = "-density 1200 -trim -transparent white -antialias +repage"
            'Execute IMconv & " " & PdfPngSwitches & " """ & FilePrefix & ".pdf"" """ & FilePrefix & ".png""", TempPath, debugMode
            
        Else
        ' dvi to png route
            LabelProcess.Caption = "LaTeX to DVI..."
            FrameProcess.Repaint
            RetVal& = Execute("pdflatex -shell-escape -output-format dvi -interaction=batchmode """ + FilePrefix + ".tex""", TempPath, debugMode, TimeOutTime)
            If (RetVal& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".dvi")) Then
                ' Error in Latex code
                ' Read log file and show it to the user
                If fs.FileExists(TempPath & FilePrefix & ".log") Then
                    Set LogFile = fs.OpenTextFile(TempPath + FilePrefix + ".log", ForReading)
                    LogFileViewer.TextBox1.Text = LogFile.ReadAll
                    LogFile.Close
                    LogFileViewer.TextBox1.ScrollBars = fmScrollBarsBoth
                    LogFileViewer.Show 1
                Else
                    MsgBox "latex did not return in " & TimeOutTimeString & " seconds and may have hung." _
                    & vbNewLine & "Please make sure your code compiles outside IguanaTex." _
                    & vbNewLine & "You may also try generating in Debug mode, as it will let you know if any font/package is missing."
                End If
                FrameProcess.Visible = False
                Exit Sub
            End If
            LabelProcess.Caption = "DVI to PNG..."
            FrameProcess.Repaint
            DviPngSwitches = "-q -D " & OutputDpiString & " -T tight"  ' monitor is 96 dpi or higher; we use OutputDpi (=1200 by default) dpi to get a crisper display, and rescale later on for new displays to match the point size
            If checkboxTransp.Value = True Then
                DviPngSwitches = DviPngSwitches & " -bg Transparent"
            End If
            ' If the user created a .png by using the standalone class with convert, we use that, else we use dvipng
            If Not fs.FileExists(TempPath & FilePrefix & ".png") Then
                RetValConv& = Execute("dvipng " & DviPngSwitches & " -o """ & FilePrefix & ".png"" """ & FilePrefix & ".dvi""", TempPath, debugMode, TimeOutTime)
                If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".png")) Then
                    MsgBox "dvipng failed, or did not return in " & TimeOutTimeString & " seconds and may have hung." _
                    & vbNewLine & "You may want to try compiling using the PDF->PNG option." _
                    & vbNewLine & "You may also try generating in Debug mode, as it will let you know if any font is missing."
                    FrameProcess.Visible = False
                    Exit Sub
                End If
            End If
        End If
        OutputType = "PNG"
        FinalFilename = FilePrefix & ".png"
    End If
    ' Latex run successful.
    
    
    ' Now we prepare the insertion of the image
    LabelProcess.Caption = "Insert image..."
    FrameProcess.Repaint
    
    ' If we are in Edit mode, store parameters of old image
    Dim PosX As Single
    Dim PosY As Single
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim oldshape As Shape
    Dim oldshapeIsEMF As Boolean
    Dim s As Shape
    IsInGroup = False
    If ButtonRun.Caption = "ReGenerate" Then
        If Sel.ShapeRange.Type = msoGroup And Sel.HasChildShapeRange Then
            ' Old image is part of a group
            Set oldshape = Sel.ChildShapeRange(1)
            IsInGroup = True
            Dim arr() As Variant ' gather all shapes to be regrouped later on
            j = 0
            For Each s In Sel.ShapeRange.GroupItems
                If s.name <> oldshape.name Then
                    j = j + 1
                    ReDim Preserve arr(1 To j)
                    arr(j) = s.name
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
            MaxGroupLevel = TagGroupHierarchy(arr, oldshape.name)
            
        Else
            Set oldshape = Sel.ShapeRange(1)
        End If
        PosX = oldshape.Left
        PosY = oldshape.Top
        oldshapeIsEMF = False
        If oldshape.Tags.Item("BitmapVector") <> "" Then
            If oldshape.Tags.Item("BitmapVector") = 1 Then
                oldshapeIsEMF = True
            End If
            'oldshapeIsEMF = oldshape.Tags.Item("EMFOUTPUT")
        'Else
            'oldshapeIsEMF = False
            'Call oldshape.Export(TempPath & FilePrefix & "_oldshape.png", ppShapeFormatPNG)
            'oldshapeDPI = GetImageFileDPI(TempPath & FilePrefix & "_oldshape.png")
        End If
    Else
        PosX = 200
        PosY = 200
    End If
            
    ' Get scaling factors
    Dim isTexpoint As Boolean
    Dim tScaleWidth As Single, tScaleHeight As Single
    'Dim RelToOrigSizeFlag As MsoTriState
    MagicScalingFactorEMF = 1 / 100 ' Magical scaling factor for EMF
    MagicScalingFactorPNG = 1 / OutputDpi
    If UseEMF Then
        MagicScalingFactor = MagicScalingFactorEMF
        'RelToOrigSizeFlag = msoFalse
    Else
        MagicScalingFactor = MagicScalingFactorPNG
        'RelToOrigSizeFlag = msoTrue
    End If
    
    If ButtonRun.Caption <> "ReGenerate" Or CheckBoxReset.Value Then
        PointSize = val(textboxSize.Text)
        tScaleWidth = PointSize / 10 * default_screen_dpi * MagicScalingFactor  ' 1/10 is for the default LaTeX point size (10 pt)
        tScaleHeight = tScaleWidth
    Else
        ' Handle the case of Texpoint displays
        isTexpoint = False
        Dim OldDpi As Long
        OldDpi = OutputDpi
        With oldshape.Tags
            If .Item("TEXPOINTSCALING") <> "" Then
                isTexpoint = True
                tScaleWidth = val(.Item("TEXPOINTSCALING")) * default_screen_dpi * MagicScalingFactor
                tScaleHeight = tScaleWidth
            End If
            If .Item("OUTPUTDPI") <> "" Then
                OldDpi = val(.Item("OUTPUTDPI"))
            End If
        End With
        'OldMagicScalingFactorPNG = highdpi_rescaling / OldDpi
        If Not isTexpoint Then ' modifying a normal display, either PNG or EMF
            HeightOld = oldshape.Height
            WidthOld = oldshape.Width
            tScaleHeight = 1
            tScaleWidth = 1
            If oldshapeIsEMF = False Then ' this deals with displays from very old versions of IguanaTex that lack proper size tags
                oldshape.ScaleHeight 1#, msoTrue
                oldshape.ScaleWidth 1#, msoTrue
                tScaleHeight = HeightOld / oldshape.Height * 960 / OutputDpi ' 0.8=960/1200 is there to preserve scaling of displays created with old versions of IguanaTex
                tScaleWidth = WidthOld / oldshape.Width * 960 / OutputDpi
            End If
            With oldshape.Tags
                If .Item("ORIGINALHEIGHT") <> "" Then
                    tmpHeight = val(.Item("ORIGINALHEIGHT"))
                    tScaleHeight = HeightOld / tmpHeight * OldDpi / OutputDpi
                End If
                If .Item("ORIGINALWIDTH") <> "" Then
                    tmpWidth = val(.Item("ORIGINALWIDTH"))
                    tScaleWidth = WidthOld / tmpWidth * OldDpi / OutputDpi
                End If
            End With
            If UseEMF = True And oldshapeIsEMF = False Then
                tScaleHeight = tScaleHeight * MagicScalingFactorEMF / MagicScalingFactorPNG
                tScaleWidth = tScaleWidth * MagicScalingFactorEMF / MagicScalingFactorPNG
            ElseIf UseEMF = False And oldshapeIsEMF = True Then
                tScaleHeight = tScaleHeight / MagicScalingFactorEMF * MagicScalingFactorPNG
                tScaleWidth = tScaleWidth / MagicScalingFactorEMF * MagicScalingFactorPNG
            End If
        End If
    End If
    
    
    ' Insert image and rescale it
    Dim newShape As Shape
    Set newShape = AddDisplayShape(TempPath + FinalFilename, PosX, PosY)
    
    If UseEMF Then
        ' Rescale the EMF picture before converting into PPT object
        Set newShape = ConvertEMF(newShape, VectorScalingX * tScaleWidth, VectorScalingY * tScaleHeight)
        ' Tag shape and its components with their "original" sizes,
        ' which we get by dividing their current height/width by the scaling factors applied above
        newShape.Tags.Add "ORIGINALHEIGHT", newShape.Height / tScaleHeight
        newShape.Tags.Add "ORIGINALWIDTH", newShape.Width / tScaleWidth
        If newShape.Type = msoGroup Then
            For Each s In newShape.GroupItems
                s.Tags.Add "ORIGINALHEIGHT", s.Height / tScaleHeight
                s.Tags.Add "ORIGINALWIDTH", s.Width / tScaleWidth
            Next
        End If
'        ' Alternative way of doing this:
'        ' Re-insert EMF picture at the original size to be able to tag each object with its "pre-rescaling" size
'        Dim unscaledShape As Shape
'        Set unscaledShape = AddDisplayShape(TempPath + FinalFilename, PosX, PosY)
'        Set unscaledShape = ConvertEMF(unscaledShape, VectorScalingX, VectorScalingY)


    Else
        ' Resize to the true size of the png file and adjust using the manual scaling factors set in Main Settings
        With newShape
            .ScaleHeight 1#, msoTrue
            .ScaleWidth 1#, msoTrue
            .LockAspectRatio = msoFalse
            .ScaleHeight BitmapScalingY, msoFalse
            .ScaleWidth BitmapScalingX, msoFalse
            .Tags.Add "OUTPUTDPI", OutputDpi ' Stores this display's resolution
            ' Add tags storing the original height and width, used next time to keep resizing ratio.
            .Tags.Add "ORIGINALHEIGHT", newShape.Height
            .Tags.Add "ORIGINALWIDTH", newShape.Width
            ' Apply scaling factors
            .ScaleHeight tScaleHeight, msoFalse
            .ScaleWidth tScaleWidth, msoFalse
            .LockAspectRatio = msoTrue
        End With
    End If
    
    ' Apply scaling factors
'    newShape.LockAspectRatio = msoFalse
'    newShape.ScaleHeight tScaleHeight, msoFalse
'    newShape.ScaleWidth tScaleWidth, msoFalse
'    newShape.LockAspectRatio = msoTrue
    
    If ButtonRun.Caption = "ReGenerate" Then ' We are editing+resetting size of an old display, we keep rotation
        newShape.Rotation = oldshape.Rotation
        If Not CheckBoxReset.Value Then
            newShape.LockAspectRatio = oldshape.LockAspectRatio ' Unlock aspect ratio if old display had it unlocked
        End If
    End If

'    If ButtonRun.Caption <> "ReGenerate" Or CheckBoxReset.Value Then
'        newShape.ScaleHeight ScaleFactor, RelToOrigSizeFlag
'        newShape.ScaleWidth ScaleFactor, RelToOrigSizeFlag
'        If ButtonRun.Caption = "ReGenerate" Then ' We are editing+resetting size of an old display, we keep rotation
'            newShape.Rotation = oldshape.Rotation
'        End If
'        newShape.LockAspectRatio = msoTrue
'    Else
'        newShape.LockAspectRatio = msoFalse
'        newShape.ScaleHeight tScaleHeight, RelToOrigSizeFlag
'        newShape.ScaleWidth tScaleWidth, RelToOrigSizeFlag
'        newShape.LockAspectRatio = oldshape.LockAspectRatio
'        newShape.Rotation = oldshape.Rotation
'    End If
    
    
    
    
    
    
    
    
    
    
    ' Add tags
    Call AddTagsToShape(newShape)
    If UseEMF = True And newShape.Type = msoGroup Then
        Set s = newShape.GroupItems(1)
        Call AddTagsToShape(s) 'only left most for now, to make things simple
        For Each s In newShape.GroupItems
        '    Call AddTagsToShape(s)
            s.Tags.Add "EMFchild", True
        Next
    End If
    
    ' Copy animation settings and formatting from old image, then delete it
    If ButtonRun.Caption = "ReGenerate" Then
        Dim TransferDesign As Boolean
        TransferDesign = True
        If UseEMF <> oldshapeIsEMF Then
            TransferDesign = False
        End If
        If IsInGroup Then
            ' Transfer format to new shape
            MatchZOrder oldshape, newShape
            If TransferDesign Then
                oldshape.PickUp
                newShape.Apply
            End If
            ' Handle the case of shape within EMF group.
            Dim DeleteLowestLayer As Boolean
            DeleteLowestLayer = False
            If oldshape.Tags.Item("EMFchild") <> "" Then
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
            arr(j) = newShape.name
            If DeleteLowestLayer Then
                Dim arr_remain() As Variant
                j_remain = 0
                For Each n In arr
                    Set s = ActiveWindow.Selection.SlideRange.Shapes(n)
                    ThisShapeLevel = 0
                    For i_tag = 1 To s.Tags.count
                        If (s.Tags.name(i_tag) = "LAYER") Then
                            ThisShapeLevel = val(s.Tags.Value(i_tag))
                        End If
                    Next
                    If ThisShapeLevel = 1 Then
                        s.Delete
                    Else
                        j_remain = j_remain + 1
                        ReDim Preserve arr_remain(1 To j_remain)
                        arr_remain(j_remain) = s.name
                    End If
                Next
                newShape.Tags.Add "LAYER", 2
                arr = arr_remain
            Else
                newShape.Tags.Add "LAYER", 1
            End If
            newShape.Tags.Add "SELECTIONNAME", newShape.name
            
            ' Hierarchically re-group elements
            For Level = 1 To MaxGroupLevel
                Dim CurrentLevelArr() As Variant
                j_current = 0
                For Each n In arr
                    ThisShapeLevel = 0
                    Dim ThisShapeSelectionName As String
                    ThisShapeSelectionName = ""
                    On Error Resume Next
                    With ActiveWindow.Selection.SlideRange.Shapes(n).Tags
                        For i_tag = 1 To .count
                            If (.name(i_tag) = "LAYER") Then
                                ThisShapeLevel = val(.Value(i_tag))
                            End If
                            If (.name(i_tag) = "SELECTIONNAME") Then
                                ThisShapeSelectionName = .Value(i_tag)
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
                    arr(j) = newGroup.name
                    newGroup.Tags.Add "SELECTIONNAME", newGroup.name
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
            MoveAnimation oldshape, newShape
            MatchZOrder oldshape, newShape
            If TransferDesign Then
                If oldshapeIsEMF And oldshape.Type = msoGroup Then
                    
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
                    For Each s In newShape.GroupItems
                        s.Apply
                    Next
                    tmpshp.Delete
                    
                    ' Now we can transfer the group formatting from the temporary shape
                    TransferGroupFormat tmpGroupEMF, newShape
                    tmpGroupEMF.Delete
                Else
                    oldshape.PickUp
                    newShape.Apply
                    oldshape.Delete
                End If
            Else
                oldshape.Delete
            End If
        End If
    End If
    
    
    ' Select the new shape
    newShape.Select
    
    
    ' Delete temp files if not in debug mode
    If debugMode = False Then fs.DeleteFile TempPath + FilePrefix + "*.*"
    
    
    
    FrameProcess.Visible = False
    Unload LatexForm
Exit Sub
   
End Sub

Private Sub AddTagsToShape(vSh As Shape)
    With vSh.Tags
        .Add "LATEXADDIN", TextBox1.Text
        .Add "IguanaTexSize", val(textboxSize.Text)
        .Add "IGUANATEXCURSOR", TextBox1.SelStart
        .Add "TRANSPARENCY", checkboxTransp.Value
        .Add "FILENAME", TextBoxFile.Text
        .Add "LATEXENGINEID", ComboBoxLaTexEngine.ListIndex
        .Add "TEMPFOLDER", TextBoxTempFolder.Text
        .Add "LATEXFORMHEIGHT", LatexForm.Height
        .Add "LATEXFORMWIDTH", LatexForm.Width
        .Add "LATEXFORMWRAP", TextBox1.WordWrap
        .Add "BitmapVector", ComboBoxBitmapVector.ListIndex
    End With
End Sub

Private Function ConvertEMF(inSh As Shape, ScalingX As Single, ScalingY As Single) As Shape
    With inSh
        .ScaleHeight 1#, msoTrue
        .ScaleWidth 1#, msoTrue
        .LockAspectRatio = msoFalse
        .ScaleHeight ScalingY, msoTrue
        .ScaleWidth ScalingX, msoTrue
        .LockAspectRatio = msoTrue
    End With
    
    Dim newShape As Shape
    ' Get current slide, it will be used to group ranges
    Dim sld As Slide
    Dim SlideIndex As Long
    SlideIndex = ActiveWindow.View.Slide.SlideIndex
    Set sld = ActivePresentation.Slides(SlideIndex)

    ' Convert EMF image to object
    Dim shr As ShapeRange
    Set shr = inSh.Ungroup
    Set shr = shr.Ungroup
    ' Clean up
    shr.Item(1).Delete
    shr.Item(2).Delete
    If shr(3).GroupItems.count > 2 Then
        Set newShape = shr(3)
    Else ' only a single freeform, so not a group
        Set newShape = shr(3).GroupItems(2)
    End If
    shr(3).GroupItems(1).Delete
    
    If newShape.Type = msoGroup Then
    
        Dim emf_arr() As Variant ' gather all shapes to be regrouped later on
        j_emf = 0
        Dim delete_arr() As Variant ' gather all shapes to be deleted later on
        j_delete = 0
        Dim s As Shape
        For Each s In newShape.GroupItems
            j_emf = j_emf + 1
            ReDim Preserve emf_arr(1 To j_emf)
            If s.Type = msoLine Then
                emf_arr(j_emf) = LineToFreeform(s).name
                j_delete = j_delete + 1
                ReDim Preserve delete_arr(1 To j_delete)
                delete_arr(j_delete) = s.name
            Else
                emf_arr(j_emf) = s.name
                s.Line.Visible = msoFalse
            End If
        Next
        newShape.Ungroup
        If j_delete > 0 Then
            sld.Shapes.Range(delete_arr).Delete
        End If
        Set newShape = sld.Shapes.Range(emf_arr).Group
    
    Else
        If newShape.Type = msoLine Then
            newShapeName = LineToFreeform(newShape).name
            newShape.Delete
            Set newShape = sld.Shapes(newShapeName)
        Else
            newShape.Line.Visible = msoFalse
        End If
    End If
    newShape.LockAspectRatio = msoTrue
    Set ConvertEMF = newShape
End Function

Private Function LineToFreeform(s As Shape) As Shape
    t = s.Line.Weight
    Dim ApplyTransform As Boolean
    ApplyTransform = True
    If s.Height = 0 Then ' Horizontal line
        x1 = s.Left
        y1 = s.Top - t / 2
        x2 = x1 + s.Width
        y2 = y1
        x3 = x2
        y3 = s.Top + t / 2
        x4 = x1
        y4 = y3
    ElseIf s.Width = 0 Then ' Vertical line
        x1 = s.Left - t / 2
        y1 = s.Top
        x2 = s.Left + t / 2
        y2 = y1
        x3 = x2
        y3 = y2 + s.Height
        x4 = x1
        y4 = y3
    Else ' Hopefully this will never happen, but we're dealing with a line that's neither horizontal nor vertical
        ApplyTransform = False
    End If
    
    
    If ApplyTransform Then
        Dim builder As FreeformBuilder
        Set builder = ActiveWindow.Selection.SlideRange(1).Shapes.BuildFreeform(msoEditingCorner, x1, y1)
        builder.AddNodes msoSegmentLine, msoEditingAuto, x2, y2
        builder.AddNodes msoSegmentLine, msoEditingAuto, x3, y3
        builder.AddNodes msoSegmentLine, msoEditingAuto, x4, y4
        builder.AddNodes msoSegmentLine, msoEditingAuto, x1, y1
        Dim oSh As Shape
        Set oSh = builder.ConvertToShape
        oSh.Fill.ForeColor = s.Line.ForeColor
        oSh.Line.Visible = msoFalse
        oSh.Rotation = s.Rotation
        Set LineToFreeform = oSh
    Else
        LineToFreeform = s
    End If
End Function

Private Function IsInArray(arr As Variant, valueToCheck As String) As Boolean
    IsInArray = False
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
        TargetGroupName = Sel.ShapeRange(1).name
        
        Dim Arr_In() As Variant ' shapes in the same group
        Dim Arr_Out() As Variant ' shapes not in the same group
        
        ' Split range according to whether elements are in the same group or not
        j_in = 0
        j_out = 0
        For Each n In arr
            ActiveWindow.Selection.SlideRange.Shapes(n).Select
            If Sel.ShapeRange.Type = msoGroup Then
                ' object is in group
                If Sel.ShapeRange(1).name = TargetGroupName Then
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
        tmp = TagGroupHierarchy(Arr_In, TargetName)
        TagGroupHierarchy = tmp + 1
        
        ' For all elements not in that group, tag them
        For Each n In Arr_Out
            ActiveWindow.Selection.SlideRange.Shapes(n).Select
            ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "LAYER", TagGroupHierarchy
            If Sel.ShapeRange.Type = msoGroup Then
                ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "SELECTIONNAME", Sel.ShapeRange(1).name
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

' Add picture as shape taking care of not inserting it in empty placeholder
Private Function AddDisplayShape(path As String, PosX As Single, PosY As Single) As Shape
' from http://www.vbaexpress.com/forum/showthread.php?47687-Addpicture-adds-the-picture-to-a-placeholder-rather-as-a-new-shape
' modified based on http://www.vbaexpress.com/forum/showthread.php?37561-Delete-empty-placeholders
    Dim oshp As Shape
    Dim osld As Slide
    On Error Resume Next
    Set osld = ActiveWindow.Selection.SlideRange(1)
    If Err <> 0 Then Exit Function
    On Error GoTo 0
    For Each oshp In osld.Shapes
        If oshp.Type = msoPlaceholder Then
            If oshp.PlaceholderFormat.ContainedType = msoAutoShape Then
                If oshp.HasTextFrame Then
                    If Not oshp.TextFrame.HasText Then oshp.TextFrame.TextRange = "DUMMY"
                End If
            End If
        End If
    Next oshp
    Set AddDisplayShape = osld.Shapes.AddPicture(path, msoFalse, msoTrue, PosX, PosY, -1, -1)
    For Each oshp In osld.Shapes
        If oshp.Type = msoPlaceholder Then
            If oshp.PlaceholderFormat.ContainedType = msoAutoShape Then
                If oshp.HasTextFrame Then
                    If oshp.TextFrame.TextRange = "DUMMY" Then oshp.TextFrame.DeleteText
                End If
            End If
        End If
    Next oshp
End Function


Private Function BoundingBoxString(BBXFile As String) As String
    Const ForReading = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtStream = fso.OpenTextFile(BBXFile, ForReading, False)
    Dim TextSplit As Variant
    Dim OutputDpiString As String
    OutputDpiString = TextBoxLocalDPI.Text
    Dim OutputDpi As Long
    OutputDpi = val(OutputDpiString)
    Do While Not txtStream.AtEndOfStream
    tmptext = txtStream.ReadLine
    TextSplit = Split(tmptext, " ")
    If TextSplit(0) = "%%HiResBoundingBox:" Then
        llx = val(TextSplit(1))
        lly = val(TextSplit(2))
        urx = val(TextSplit(3))
        ury = val(TextSplit(4))
        'compute size and offset
        sx = CStr(Round((urx - llx) / 72 * OutputDpi))
        sy = CStr(Round((ury - lly) / 72 * OutputDpi))
        cx = Str(-llx)
        cy = Str(-lly)
    End If
    Loop
    txtStream.Close
    BoundingBoxString = " -g" & sx & "x" & sy & " -c ""<</Install {" & cx & " " & cy & " translate}>> setpagedevice"""
End Function

Private Sub SaveSettings()
    RegPath = "Software\IguanaTex"
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Transparent", REG_DWORD, BoolToInt(checkboxTransp.Value)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Debug", REG_DWORD, BoolToInt(checkboxDebug.Value)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "PointSize", REG_DWORD, CLng(val(textboxSize.Text))
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LatexCode", REG_SZ, CStr(TextBox1.Text)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LatexCodeCursor", REG_DWORD, CLng(TextBox1.SelStart)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LatexFormHeight", REG_DWORD, CLng(LatexForm.Height)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LatexFormWidth", REG_DWORD, CLng(LatexForm.Width)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Multipage", REG_SZ, MultiPage1.Value
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LatexFormWrap", REG_DWORD, BoolToInt(TextBox1.WordWrap)
    'SetRegistryValue HKEY_CURRENT_USER, RegPath, "EMFoutput", REG_DWORD, BoolToInt(CheckBoxEMF.Value)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "BitmapVector", REG_DWORD, ComboBoxBitmapVector.ListIndex
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "OutputDpi", REG_DWORD, CLng(val(TextBoxLocalDPI.Text))
    
    
End Sub

Private Sub LoadSettings()
    RegPath = "Software\IguanaTex"
    checkboxTransp.Value = CBool(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Transparent", True))
    checkboxDebug.Value = CBool(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Debug", False))
    textboxSize.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "PointSize", "20")
    TextBox1.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexCode", "\documentclass{article}" & Chr(13) & "\usepackage{amsmath}" & Chr(13) & "\pagestyle{empty}" & Chr(13) & "\begin{document}" & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "\end{document}")
    TextBox1.SelStart = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexCodeCursor", 0)
    MultiPage1.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Multipage", 0)
    TextBox1.Font.Size = val(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "EditorFontSize", "10"))
    TextBoxTempFolder.Text = GetTempPath()
    TextBox1.WordWrap = CBool(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexFormWrap", True))
    ToggleButtonWrap.Value = TextBox1.WordWrap
    
    LaTexEngineList = Array("pdflatex", "pdflatex", "xelatex", "lualatex", "platex")
    LaTexEngineDisplayList = Array("latex (DVI)", "pdflatex", "xelatex", "lualatex", "platex")
    UsePDFList = Array(False, True, True, True, True)
    ComboBoxLaTexEngine.List = LaTexEngineDisplayList
    ComboBoxLaTexEngine.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LaTeXEngineID", 0)
    TextBoxLocalDPI.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "OutputDpi", "1200")
    ComboBoxBitmapVector.List = Array("Bitmap", "Vector")
    ComboBoxBitmapVector.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "BitmapVector", 0)
            
    TemplateSortedListString = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TemplateSortedList", "0")
    TemplateSortedList = UnpackStringToArray(TemplateSortedListString)
    TemplateNameSortedListString = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TemplateNameSortedList", "New Template")
    ComboBoxTemplate.List = UnpackStringToArray(TemplateNameSortedListString)
End Sub

Private Function BoolToInt(val) As Long
    If val Then
        BoolToInt = 1&
    Else
        BoolToInt = 0&
    End If
End Function

Private Sub ButtonTexPath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker) 'msoFileDialogFolderPicker
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = TextBoxFile.Text
    fd.Filters.Clear
    fd.Filters.Add "Tex Files", "*.tex", 1
    
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            TextBoxFile.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If

    Set fd = Nothing
    TextBoxFile.SetFocus
End Sub

Private Sub ComboBoxBitmapVector_Change()
    Apply_BitmapVector_Change
End Sub

Private Sub Apply_BitmapVector_Change()
    If ComboBoxBitmapVector.ListIndex = 1 Then
        checkboxTransp.Enabled = False
        checkboxTransp.Value = True
        TextBoxLocalDPI.Enabled = False
        LabelDPI.Enabled = False
    Else
        checkboxTransp.Enabled = True
        TextBoxLocalDPI.Enabled = True
        LabelDPI.Enabled = True
    End If

End Sub

Private Sub CheckBoxReset_Click()
    If CheckBoxReset.Value = True Then
        textboxSize.Enabled = True
    Else
        textboxSize.Enabled = False
    End If
End Sub

Private Sub ButtonAbout_Click()
    AboutBox.Show 1
End Sub


Private Sub ButtonMakeDefault_Click()
    SaveSettings
    Select Case MultiPage1.Value
        Case 0 ' Direct input
            TextBox1.SetFocus
        Case 1 ' Read from file
            TextBoxFile.SetFocus
        Case Else ' Templates
            TextBoxTemplateCode.SetFocus
    End Select
End Sub

Private Sub CmdButtonExternalEditor_Click()
        
    ' Put the temporary path in the right format
    If Right(TextBoxTempFolder.Text, 1) <> "\" Then
        TextBoxTempFolder.Text = TextBoxTempFolder.Text & "\"
    End If
    Dim TempPath As String
    TempPath = TextBoxTempFolder.Text
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
    
    Dim FilePrefix As String
    FilePrefix = GetFilePrefix()
    
    ' Test if path writable
    If Not IsPathWritable(TempPath) Then
        MsgBox "The temporary folder " & TempPath & " appears not to be writable."
        Exit Sub
    End If
    
    ' Write latex to a temp file
    Call WriteLaTeX2File(TempPath, FilePrefix)
    
    ' Launch external editor
    On Error GoTo ShellError
    Shell """" & GetEditorPath() & """ """ & TempPath & FilePrefix & ".tex""", vbNormalFocus
    
    ' Show dialog form to reload from file or cancel
    ExternalEditorForm.Show
    Exit Sub
    
ShellError:
    MsgBox "Error Launching External Editor." & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
    Exit Sub
End Sub

Private Sub CmdButtonImportCode_Click()
    TextBoxTemplateCode.Text = TextBox1.Text
    TextBoxTemplateCode.SelStart = TextBox1.SelStart
    TextBoxTemplateCode.SetFocus
End Sub

Private Sub CmdButtonLoadTemplate_Click()
    If TextBoxTemplateCode.Text = "" Then
        MsgBox "Please select a template to be loaded"
    Else
        TextBox1.Text = TextBoxTemplateCode.Text
        TextBox1.SelStart = TextBoxTemplateCode.SelStart
        MultiPage1.Value = 0
        Call ToggleInputMode
    End If
End Sub

Private Sub CmdButtonRemoveTemplate_Click()
    Dim RemovedIndex As Long
    RemovedIndex = ComboBoxTemplate.ListIndex
    If ComboBoxTemplate.ListCount > 1 Then
        ' We should also be deleting the registry entry, but well, it does not take much space and will likely get reused anyway
        ComboBoxTemplate.RemoveItem RemovedIndex
        
        ' update the array that contains the sorted list of template IDs
        tmpID = TemplateSortedList(RemovedIndex)
        Dim i As Long
        For i = RemovedIndex To UBound(TemplateSortedList) - 1
            TemplateSortedList(i) = TemplateSortedList(i + 1)
        Next i
        TemplateSortedList(UBound(TemplateSortedList)) = tmpID
        'NumberOfTemplates = NumberOfTemplates - 1
    Else
        ComboBoxTemplate.Clear
        ComboBoxTemplate.AddItem "New Template" 'prepare spot for new template
        ComboBoxTemplate.Text = ""
        'NumberOfTemplates = 1
    End If
    Call UpdateTemplateRegistry
    ComboBoxTemplate.ListIndex = RemovedIndex
    TextBoxTemplateCode.SetFocus
End Sub

Private Sub CmdButtonSaveTemplate_Click()
    ' get the right ID from the array of sorted template IDs
    templateID = TemplateSortedList(ComboBoxTemplate.ListIndex)
    ' add trailing new line if there isn't one: this helps with a bug where text with multi-byte characters gets chopped
    If Not Right(TextBoxTemplateCode.Text, 1) = Chr(13) And Not Right(TextBoxTemplateCode.Text, 1) = Chr(10) Then
        TextBoxTemplateCode.Text = TextBoxTemplateCode.Text & Chr(13)
    End If
    ' build the corresponding registry key string
    ' Save name, code, and LaTeXEngineID
    RegPath = "Software\IguanaTex"
    Dim RegStr As String
    RegStr = "TemplateCode" & templateID
    SetRegistryValue HKEY_CURRENT_USER, RegPath, RegStr, REG_SZ, CStr(TextBoxTemplateCode.Text)
    RegStr = "TemplateCodeSelStart" & templateID
    SetRegistryValue HKEY_CURRENT_USER, RegPath, RegStr, REG_DWORD, CLng(TextBoxTemplateCode.SelStart)
    RegStr = "TemplateLaTeXEngineID" & templateID
    SetRegistryValue HKEY_CURRENT_USER, RegPath, RegStr, REG_DWORD, ComboBoxLaTexEngine.ListIndex
    RegStr = "TemplateBitmapVector" & templateID
    SetRegistryValue HKEY_CURRENT_USER, RegPath, RegStr, REG_DWORD, ComboBoxBitmapVector.ListIndex
    RegStr = "TemplateTempFolder" & templateID
    SetRegistryValue HKEY_CURRENT_USER, RegPath, RegStr, REG_SZ, CStr(TextBoxTempFolder.Text)
    RegStr = "TemplateDPI" & templateID
    SetRegistryValue HKEY_CURRENT_USER, RegPath, RegStr, REG_SZ, CStr(TextBoxLocalDPI.Text)
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
    Call UpdateTemplateRegistry
    TextBoxTemplateCode.SetFocus
End Sub



Private Sub ComboBoxTemplate_Click()
    TextBoxTemplateName.Text = ComboBoxTemplate.Text
    ' Except for the empty "New Template" slot, get the code and LaTeXEngineID setting from registry
    If ComboBoxTemplate.ListIndex = ComboBoxTemplate.ListCount - 1 Then
        TextBoxTemplateCode.Text = ""
        ComboBoxLaTexEngine.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LaTeXEngineID", 0)
        TextBoxTempFolder.Text = GetTempPath()
    Else
        ' get the right ID from the array of sorted template IDs
        templateID = TemplateSortedList(ComboBoxTemplate.ListIndex)
        ' build the corresponding registry key string
        RegPath = "Software\IguanaTex"
        Dim RegStr As String
        RegStr = "TemplateCode" & templateID
        TextBoxTemplateCode.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, RegStr, "")
        RegStr = "TemplateCodeSelStart" & templateID
        TextBoxTemplateCode.SelStart = GetRegistryValue(HKEY_CURRENT_USER, RegPath, RegStr, 0)
        RegStr = "TemplateLaTeXEngineID" & templateID
        ComboBoxLaTexEngine.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, RegStr, GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LaTeXEngineID", 0))
        RegStr = "TemplateBitmapVector" & templateID
        ComboBoxBitmapVector.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, RegStr, GetRegistryValue(HKEY_CURRENT_USER, RegPath, "BitmapVector", False))
        RegStr = "TemplateTempFolder" & templateID
        TextBoxTempFolder.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, RegStr, GetTempPath())
        RegStr = "TemplateDPI" & templateID
        TextBoxLocalDPI.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, RegStr, "")
        Apply_BitmapVector_Change
    End If
    TextBoxTemplateCode.SetFocus
End Sub

Private Sub UpdateTemplateRegistry()
    ' update the list of saved templates names in the registry (will be used to initialize combo box content)
    TemplateSortedListString = PackArrayToString(TemplateSortedList)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "TemplateSortedList", REG_SZ, CStr(TemplateSortedListString)
    ' save list of template names to registry
    Dim myArray() As String
    ReDim myArray(0 To ComboBoxTemplate.ListCount - 1) As String
    Dim i As Long
    For i = LBound(myArray) To UBound(myArray)
        myArray(i) = ComboBoxTemplate.List(i)
    Next i
    TemplateNameSortedListString = PackArrayToString(myArray)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "TemplateNameSortedList", REG_SZ, CStr(TemplateNameSortedListString)
End Sub

Private Sub CmdButtonTemplateFontDown_Click()
    If TextBoxTemplateCode.Font.Size > 4 Then
        TextBoxTemplateCode.Font.Size = TextBoxTemplateCode.Font.Size - 1
    End If
End Sub

Private Sub CmdButtonTemplateFontUp_Click()
    If TextBoxTemplateCode.Font.Size < 72 Then
        TextBoxTemplateCode.Font.Size = TextBoxTemplateCode.Font.Size + 1
    End If
End Sub

Private Sub CmdButtonEditorFontDown_Click()
    If TextBox1.Font.Size > 4 Then
        TextBox1.Font.Size = TextBox1.Font.Size - 1
    End If
End Sub

Private Sub CmdButtonEditorFontUp_Click()
    If TextBox1.Font.Size < 72 Then
        TextBox1.Font.Size = TextBox1.Font.Size + 1
    End If
End Sub

Private Sub ToggleButtonWrap_Click()
    If ToggleButtonWrap.Value = True Then
        TextBox1.WordWrap = True
    Else
        TextBox1.WordWrap = False
    End If
End Sub

Private Sub UserForm_Initialize()
    LoadSettings
    
    ' With multiple monitors, the "CenterOwner" option to open the UserForm in the center of the parent window
    ' does not seem to work, at least in Office 2010.
    ' The following code to manually place the UserForm somehow makes the "CenterOwner" option work.
    ' Remark: if used with the Manual placement option, it would place the window to the left, under the ribbon.
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    
    LatexForm.textboxSize.Visible = True
    LatexForm.Label2.Visible = True
    LatexForm.Label3.Visible = True

    FrameProcess.Visible = False
    
    
End Sub

Private Sub UserForm_Activate()
    'Execute macro to enable resizeability
    MakeFormResizable
    
    RegPath = "Software\IguanaTex"
    If Not FormHeightWidthSet Then
        LatexForm.Height = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexFormHeight", 312)
        LatexForm.Width = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexFormWidth", 380)
    End If
    ResizeForm
    
End Sub

Sub RetrieveOldShapeInfo(oldshape As Shape, mainText As String)
    CheckBoxReset.Visible = True
    CheckBoxReset.Value = False
    Label2.Caption = "Reset size:"
    ButtonRun.Caption = "ReGenerate"
    ButtonRun.Accelerator = "G"
    
    TextBox1.Text = mainText
    CursorPosition = Len(TextBox1.Text)
                
    Dim FormHeightSet As Boolean
    Dim FormWidthSet As Boolean
    FormHeightSet = False
    FormWidthSet = False
     
    With oldshape.Tags
        If .Item("IGUANATEXSIZE") <> "" Then
            textboxSize.Text = .Item("IGUANATEXSIZE")
        End If
        If .Item("OUTPUTDPI") <> "" Then
            TextBoxLocalDPI.Text = .Item("OUTPUTDPI")
        End If
        If .Item("BitmapVector") <> "" Then
            ComboBoxBitmapVector.ListIndex = .Item("BitmapVector")
        End If
        If .Item("TRANSPARENCY") <> "" Then
            checkboxTransp.Value = SanitizeBoolean(.Item("TRANSPARENCY"), True)
        ElseIf .Item("TRANSPARENT") <> "" Then
            checkboxTransp.Value = SanitizeBoolean(.Item("TRANSPARENT"), True)
        End If
        If .Item("IGUANATEXCURSOR") <> "" Then
            CursorPosition = .Item("IGUANATEXCURSOR")
        End If
        If .Item("LATEXENGINEID") <> "" Then
            ComboBoxLaTexEngine.ListIndex = .Item("LATEXENGINEID")
        End If
        If .Item("LATEXFORMHEIGHT") <> "" Then
            LatexForm.Height = .Item("LATEXFORMHEIGHT")
            FormHeightSet = True
        End If
        If .Item("LATEXFORMWIDTH") <> "" Then
            LatexForm.Width = .Item("LATEXFORMWIDTH")
            FormWidthSet = True
        End If
        If .Item("LATEXFORMWRAP") <> "" Then
            TextBox1.WordWrap = SanitizeBoolean(.Item("LATEXFORMWRAP"), True)
            ToggleButtonWrap.Value = TextBox1.WordWrap
        End If
    End With
    Apply_BitmapVector_Change
    FormHeightWidthSet = FormHeightSet And FormWidthSet
    TextBox1.SelStart = CursorPosition
    textboxSize.Enabled = False
End Sub


Private Function SanitizeBoolean(Str As String, Def As Boolean) As Boolean
    On Error GoTo ErrWrongBoolean:
    SanitizeBoolean = CBool(Str)
    Exit Function
ErrWrongBoolean:
    SanitizeBoolean = Def
    Resume Next
End Function

Private Sub UserForm_Resize()
    ' Minimal size
    Select Case MultiPage1.Value
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
    checkboxTransp.Top = CheckBoxReset.Top + checkboxTransp.Height + 2
    checkboxDebug.Top = checkboxTransp.Top + checkboxTransp.Height + 2
    
    
End Sub

Private Sub MoveAnimation(oldshape As Shape, newShape As Shape)
    ' Move the animation settings of oldShape to newShape
    With ActiveWindow.Selection.SlideRange(1).TimeLine
        Dim eff As Effect
        For Each eff In .MainSequence
            If eff.Shape.name = oldshape.name Then eff.Shape = newShape
        Next
    End With
End Sub

Private Sub MatchZOrder(oldshape As Shape, newShape As Shape)
    ' Make the Z order of newShape equal to 1 higher than that of oldShape
    newShape.ZOrder msoBringToFront
    While (newShape.ZOrderPosition > oldshape.ZOrderPosition + 1)
        newShape.ZOrder msoSendBackward
    Wend
End Sub

Private Sub DeleteAnimation(oldshape As Shape)
    ' Delete the animation settings of oldShape
    With ActiveWindow.Selection.SlideRange(1).TimeLine
        For i = .MainSequence.count To 1 Step -1
            Dim eff As Effect
            Set eff = .MainSequence(i)
            If eff.Shape.name = oldshape.name Then eff.Delete
        Next
    End With
End Sub

Private Sub TransferGroupFormat(oldshape As Shape, newShape As Shape)
    On Error Resume Next
    ' Transfer group formatting
    If oldshape.Glow.Radius > 0 Then
        newShape.Glow.Color = oldshape.Glow.Color
        newShape.Glow.Radius = oldshape.Glow.Radius
        newShape.Glow.Transparency = oldshape.Glow.Transparency
    End If
    If oldshape.Reflection.Type <> msoReflectionTypeNone Then
        newShape.Reflection.Blur = oldshape.Reflection.Blur
        newShape.Reflection.Offset = oldshape.Reflection.Offset
        newShape.Reflection.Size = oldshape.Reflection.Size
        newShape.Reflection.Transparency = oldshape.Reflection.Transparency
        newShape.Reflection.Type = oldshape.Reflection.Type
    End If
    
    If oldshape.SoftEdge.Type <> msoSoftEdgeTypeNone Then
        newShape.SoftEdge.Radius = oldshape.SoftEdge.Radius
    End If
    
    If oldshape.Shadow.Visible Then
        newShape.Shadow.Visible = oldshape.Shadow.Visible
        newShape.Shadow.Blur = oldshape.Shadow.Blur
        newShape.Shadow.ForeColor = oldshape.Shadow.ForeColor
        newShape.Shadow.OffsetX = oldshape.Shadow.OffsetX
        newShape.Shadow.OffsetY = oldshape.Shadow.OffsetY
        newShape.Shadow.RotateWithShape = oldshape.Shadow.RotateWithShape
        newShape.Shadow.Size = oldshape.Shadow.Size
        newShape.Shadow.Style = oldshape.Shadow.Style
        newShape.Shadow.Transparency = oldshape.Shadow.Transparency
        newShape.Shadow.Type = oldshape.Shadow.Type
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
        newShape.ThreeD.Visible = oldshape.ThreeD.Visible
        newShape.ThreeD.Perspective = oldshape.ThreeD.Perspective
        newShape.ThreeD.FieldOfView = oldshape.ThreeD.FieldOfView
        newShape.ThreeD.LightAngle = oldshape.ThreeD.LightAngle
        'newShape.ThreeD.ProjectText = oldshape.ThreeD.ProjectText
        'If oldshape.ThreeD.PresetExtrusionDirection <> msoPresetExtrusionDirectionMixed Then
        '    newShape.ThreeD.SetExtrusionDirection oldshape.ThreeD.PresetExtrusionDirection
        'End If
        newShape.ThreeD.PresetLighting = oldshape.ThreeD.PresetLighting
        If oldshape.ThreeD.PresetLightingDirection <> msoPresetLightingDirectionMixed Then
            newShape.ThreeD.PresetLightingDirection = oldshape.ThreeD.PresetLightingDirection
        End If
        If oldshape.ThreeD.PresetLightingSoftness <> msoPresetLightingSoftnessMixed Then
            newShape.ThreeD.PresetLightingSoftness = oldshape.ThreeD.PresetLightingSoftness
        End If
        If oldshape.ThreeD.PresetMaterial <> msoPresetMaterialMixed Then
            newShape.ThreeD.PresetMaterial = oldshape.ThreeD.PresetMaterial
        End If
        If oldshape.ThreeD.PresetCamera <> msoPresetCameraMixed Then
            newShape.ThreeD.SetPresetCamera oldshape.ThreeD.PresetCamera
        End If
        newShape.ThreeD.RotationX = oldshape.ThreeD.RotationX
        newShape.ThreeD.RotationY = oldshape.ThreeD.RotationY
        newShape.ThreeD.RotationZ = oldshape.ThreeD.RotationZ
        'newShape.ThreeD.Z = oldshape.ThreeD.Z
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If CloseMode = vbFormControlMenu Then
        ' Cancel = True
        ' ButtonCancel_Click
    ' End If
End Sub


Private Function isTex(file As String)
    ext = Right$(file, Len(file) - InStrRev(file, "."))
    If ext = "tex" Then
        isTex = True
    Else
        isTex = False
    End If
End Function


Private Sub MultiPage1_Change()
    Call ToggleInputMode
End Sub


Private Sub TextBoxFile_Change()
    Set fs = CreateObject("Scripting.FileSystemObject")
    ButtonLoadFile.Enabled = fs.FileExists(TextBoxFile.Text) And isTex(TextBoxFile.Text)
End Sub

Private Sub ButtonLoadFile_Click()
    MultiPage1.Value = 0
    Call LoadTexFile
    Call ToggleInputMode
End Sub

Private Sub LoadTexFile()

    Dim fs
    Dim TexFile As Object
    Const ForReading = 1, ForWriting = 2, ForAppending = 3

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(TextBoxFile.Text) Then
        Set TexFile = fs.OpenTextFile(TextBoxFile.Text, ForReading)
        TextBox1.Text = TexFile.ReadAll
        TexFile.Close
    End If
    
End Sub

Private Sub ToggleInputMode()
    Set fs = CreateObject("Scripting.FileSystemObject")
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    
    Select Case MultiPage1.Value
        Case 0 ' Direct input
            TextBox1.SetFocus
        Case 1 ' Read from file
            TextBoxFile.SetFocus
            ButtonLoadFile.Enabled = fs.FileExists(TextBoxFile.Text) And isTex(TextBoxFile.Text)
        Case Else ' Templates
            If TextBoxTemplateName.Text = "" Then
                TextBoxTemplateName.Text = ComboBoxTemplate.Text
            End If
            TextBoxTemplateCode.SetFocus
            
    End Select
    Call UserForm_Resize
    
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
'    If imgFile.HorizontalResolution <> "" Then
'        GetImageFileDPI = Round(imgFile.HorizontalResolution)
'    End If
'End Function


