VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LatexForm 
   Caption         =   "IguanaTex"
   ClientHeight    =   5880
   ClientLeft      =   21
   ClientTop       =   336
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

Dim theAppEventHandler As New AppEventHandler

Sub InitializeApp()
    Set theAppEventHandler.App = Application
    
    AddMenuItem "New Latex e&quation...", "NewLatexEquation", 18 '226
    AddMenuItem "Edit Latex equation...", "EditLatexEquation", 37
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
    
    RemoveMenuItem "New Latex e&quation..."
    RemoveMenuItem "Edit Latex equation..."
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


Sub ButtonRun_Click()
    Dim TempPath As String
    TempPath = GetTempPath()
    FilePrefix = GetFilePrefix()
    
    Dim debugMode As Boolean
    If checkboxDebug.Value Then
        debugMode = True
    Else
        debugMode = False
    End If
    
    ' Read settings
    RegPath = "Software\IguanaTex"
    Dim UseUTF8 As Boolean
    Dim UsePDF As Boolean
    UseUTF8 = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "UseUTF8", True)
    'UsePDF = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "UsePDF", False)
    gs_command = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "GS Command", "C:\Program Files (x86)\gs\gs9.15\bin\gswin32c.exe")
    IMconv = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "IMconv", "C:\Program Files\ImageMagick\convert.exe")
    'tex2pdf_command = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LaTeXEngine", "pdflatex")
    LaTeXEngineID = ComboBoxLaTexEngine.ListIndex
    tex2pdf_command = LaTexEngineList(LaTeXEngineID)
    UsePDF = UsePDFList(LaTeXEngineID)
    
    Dim TimeOutTimeString As String
    Dim TimeOutTime As Long
    TimeOutTimeString = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TimeOutTime", "20") ' Wait 20 seconds for the processes to complete
    TimeOutTime = val(TimeOutTimeString) * 1000
    
    ' Read current dpi in: this will be used when rescaling and optionally in pdf->png conversion
    dpi = lDotsPerInch
    highdpi_rescaling = 1 ' will be used to account for dvipng's handling of high-dpi displays
        
    ' Write latex to a temp file
    Set fs = CreateObject("Scripting.FileSystemObject")
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
    
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    
    If fs.FileExists(TempPath & FilePrefix & ".png") Then
        fs.DeleteFile TempPath + FilePrefix + "*.*" 'Make sure we don't keep old files
    End If
    
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
    
    Dim LogFile As Object
            
    ' Run latex
    FrameProcess.Visible = True
    
    If UsePDF = True Then
    ' pdf to png route
        LabelProcess.Caption = "LaTeX to PDF..."
        FrameProcess.Repaint
        RetVal& = Execute("""" & tex2pdf_command & """ -shell-escape -interaction=batchmode """ + FilePrefix + ".tex""", TempPath, debugMode, TimeOutTime)
            
        If (RetVal& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".pdf")) Then
            ' Error in Latex code
            ' Read log file and show it to the user
            If fs.FileExists(TempPath & FilePrefix & ".log") Then
                Set LogFile = fs.OpenTextFile(TempPath + FilePrefix + ".log", ForReading)
                LogFileViewer.TextBox1.Text = LogFile.ReadAll
                LogFile.Close
                LogFileViewer.TextBox1.ScrollBars = fmScrollBarsBoth
                LogFileViewer.Show 1
            Else
                MsgBox tex2pdf_command & " did not return in " & TimeOutTimeString & " seconds and may have hung." & vbNewLine & "Please make sure your code compiles outside IguanaTex."
            End If
            FrameProcess.Visible = False
            Exit Sub
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
        RetValConv& = Execute("""" & gs_command & """ -q -dBATCH -dNOPAUSE -sDEVICE=pngalpha -r1200 -sOutputFile=""" & FilePrefix & "_tmp.png""" & BBString & " -f """ & TempPath & FilePrefix & ".pdf""", TempPath, debugMode, TimeOutTime)
        If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & "_tmp.png")) Then
            ' Error in PDF to PNG conversion
            MsgBox "Error while using Ghostscript to convert from PDF to PNG. Is your path correct?"
            FrameProcess.Visible = False
            Exit Sub
        End If
        ' Unfortunately, the resulting file has a metadata DPI of 1200, not the default screen one (usually 96),
        ' so there is a discrepancy with the dvipng output.
        ' The only workaround I have found so far is to use Imagemagick's convert to change the DPI (but not the pixel size!)
        ' Execute """" & IMconv & """ -units PixelsPerInch """ & FilePrefix & "_tmp.png"" -density " & CStr(dpi) & " """ & FilePrefix & ".png""", TempPath, debugMode
        RetValConv& = Execute("""" & IMconv & """ -units PixelsPerInch """ & FilePrefix & "_tmp.png"" -density " & CStr(dpi) & " """ & FilePrefix & ".png""", TempPath, debugMode, TimeOutTime)
        If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".png")) Then
            ' Error in PDF to PNG conversion
            MsgBox "Error while using ImageMagick to change the PNG DPI. Is your path correct?" & vbNewLine & "The full path is needed to avoid conflict with Windows's built-in convert.exe."
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
                MsgBox "pdflatex did not return in " & TimeOutTimeString & " seconds and may have hung." & vbNewLine & "Please make sure your code compiles outside IguanaTex."
            End If
            FrameProcess.Visible = False
            Exit Sub
        End If
        LabelProcess.Caption = "DVI to PNG..."
        FrameProcess.Repaint
        DviPngSwitches = "-q -D 1200 -T tight"  ' monitor is 96 dpi or higher; we use 1200 dpi to get a crisper display, and rescale later on for new displays to match the point size
        If checkboxTransp.Value = True Then
            DviPngSwitches = DviPngSwitches & " -bg Transparent"
        End If
        ' If the user created a .png by using the standalone class with convert, we use that, else we use dvipng
        If Not fs.FileExists(TempPath & FilePrefix & ".png") Then
            RetValConv& = Execute("dvipng " & DviPngSwitches & " -o """ & FilePrefix & ".png"" """ & FilePrefix & ".dvi""", TempPath, debugMode, TimeOutTime)
            If (RetValConv& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".png")) Then
                MsgBox "dvipng failed, or did not return in " & TimeOutTimeString & " seconds and may have hung." & vbNewLine & "You may want to try compiling using the PDF->PNG option."
                FrameProcess.Visible = False
                Exit Sub
            End If
        End If
        highdpi_rescaling = 96 / dpi
    End If
    
    FinalFilename = FilePrefix & ".png"
    
    LabelProcess.Caption = "Insert image..."
    FrameProcess.Repaint
    ' Latex run successful.
    ' If we are in Edit mode, store parameters of old image
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim oldShape As Shape
    IsInGroup = False
    If ButtonRun.Caption = "ReGenerate" Then
        If Sel.ShapeRange.Type = msoGroup Then
            Set oldShape = Sel.ChildShapeRange(1)
            IsInGroup = True
            Dim arr() As Variant ' gather all shapes to be regrouped later on
            j = 0
            Dim s As Shape
            For Each s In Sel.ShapeRange.GroupItems
                If s.name <> oldShape.name Then
                    j = j + 1
                    ReDim Preserve arr(1 To j)
                    arr(j) = s.name
                End If
            Next
            ' Store the group's animation and Zorder info in a dummy object tmpGroup
            Dim oldShapeRange As ShapeRange
            Set oldShapeRange = Sel.ShapeRange
            Dim oldGroup As Shape
            Set oldGroup = oldShapeRange(1)
            Dim tmpGroup As Shape
            Set tmpGroup = ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeDiamond, 1, 1, 1, 1)
            MoveAnimation oldGroup, tmpGroup
            MatchZOrder oldGroup, tmpGroup
            
            ' Tag all elements in the group with their hierarchy level and their name or group name
            Dim MaxGroupLevel As Long
            MaxGroupLevel = TagGroupHierarchy(arr, oldShape.name)
            
        Else
            Set oldShape = Sel.ShapeRange(1)
        End If
        posX = oldShape.Left
        posY = oldShape.Top
    Else
        posX = 200
        posY = 200
    End If
    
    ' Insert image
    Dim newShape As Shape
    Set newShape = ActiveWindow.Selection.SlideRange.Shapes.AddPicture(TempPath + FinalFilename, msoFalse, msoTrue, posX, posY, -1, -1)
    ' Resize to the true size of the png file
    newShape.ScaleHeight 1#, msoTrue
    newShape.ScaleWidth 1#, msoTrue
    ' Add tags storing the original height and width, used next time to keep resizing ratio.
    newShape.Tags.Add "ORIGINALHEIGHT", newShape.Height
    newShape.Tags.Add "ORIGINALWIDTH", newShape.Width
    
    
    
    
    
    ' Scale it
    If ButtonRun.Caption <> "ReGenerate" Or CheckBoxReset.Value = True Then
        PointSize = val(textboxSize.Text)
        ScaleFactor = PointSize / 10 * dpi / 1200 * highdpi_rescaling  ' 1/10 is for the default LaTeX point size (10 pt)
        newShape.ScaleHeight ScaleFactor, msoTrue
        newShape.ScaleWidth ScaleFactor, msoTrue
        If ButtonRun.Caption = "ReGenerate" Then ' We are editing+resetting size of an old display, we keep rotation
            newShape.Rotation = oldShape.Rotation
        End If
    Else
        ' Handle the case of Texpoint displays
        Dim isTexpoint As Boolean
        isTexpoint = False
        With oldShape.Tags
            For i = 1 To .count
                If (.name(i) = "TEXPOINTSCALING") Then
                    isTexpoint = True
                    tScaleWidth = val(.Value(i)) * dpi / 1200 * highdpi_rescaling
                End If
            Next
        End With
        If isTexpoint Then
            newShape.LockAspectRatio = msoTrue
            newShape.ScaleWidth tScaleWidth, msoTrue
            newShape.LockAspectRatio = oldShape.LockAspectRatio
            newShape.Rotation = oldShape.Rotation
        Else ' modifying a normal
            HeightOld = oldShape.Height
            WidthOld = oldShape.Width
            oldShape.ScaleHeight 1#, msoTrue
            oldShape.ScaleWidth 1#, msoTrue
            tScaleHeight = HeightOld / oldShape.Height * 0.8 ' 0.8=960/1200 is there to preserve scaling of displays created with old versions of IguanaTex
            tScaleWidth = WidthOld / oldShape.Width * 0.8
            With oldShape.Tags
                For i = 1 To .count
                    If (.name(i) = "ORIGINALHEIGHT") Then
                        tmpHeight = val(.Value(i))
                        tScaleHeight = HeightOld / tmpHeight
                    End If
                    If (.name(i) = "ORIGINALWIDTH") Then
                        tmpWidth = val(.Value(i))
                        tScaleWidth = WidthOld / tmpWidth
                    End If
                Next
            End With
                        
            newShape.LockAspectRatio = msoFalse
            newShape.ScaleHeight tScaleHeight, msoTrue
            newShape.ScaleWidth tScaleWidth, msoTrue
            newShape.LockAspectRatio = oldShape.LockAspectRatio
            newShape.Rotation = oldShape.Rotation
        End If
    End If
    
    ' Add tags
    newShape.Tags.Add "LATEXADDIN", TextBox1.Text
    newShape.Tags.Add "IguanaTexSize", val(textboxSize.Text)
    newShape.Tags.Add "IGUANATEXCURSOR", TextBox1.SelStart
    newShape.Tags.Add "TRANSPARENCY", checkboxTransp.Value
    newShape.Tags.Add "FILENAME", TextBoxFile.Text
    newShape.Tags.Add "INPUTTYPE", BoolToInt(MultiPage1.Value)
    newShape.Tags.Add "LATEXENGINEID", LaTeXEngineID
    
    ' Copy animation settings and formatting from old image, then delete it
    If ButtonRun.Caption = "ReGenerate" Then
        If IsInGroup Then
            ' Transfer format to new shape
            MatchZOrder oldShape, newShape
            oldShape.PickUp
            newShape.Apply
            oldShape.Delete
            
            ' Get current slide, it will be used to group ranges
            Dim sld As Slide
            Dim SlideIndex As Long
            SlideIndex = ActiveWindow.View.Slide.SlideIndex
            Set sld = ActivePresentation.Slides(SlideIndex)
            Dim newGroup As Shape
            
            ' Group all non-modified elements from old group, plus modified element
            j = j + 1
            ReDim Preserve arr(1 To j)
            arr(j) = newShape.name
            newShape.Tags.Add "LAYER", 1
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
            MoveAnimation oldShape, newShape
            MatchZOrder oldShape, newShape
            oldShape.PickUp
            newShape.Apply
            oldShape.Delete
        End If
    End If
    
    ' Select the new shape
    newShape.Select
    
    ' Delete temp files if not in debug mode
    If debugMode = False Then fs.DeleteFile TempPath + FilePrefix + "*.*"
    FrameProcess.Visible = False
    Unload LatexForm
Exit Sub

TempFolderNotWritable:
    'Debug.Print "The temporary folder " & TempPath & " appears not to be writable."
    MsgBox "The temporary folder " & TempPath & " appears not to be writable."
    
End Sub

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
        Tmp = TagGroupHierarchy(Arr_In, TargetName)
        TagGroupHierarchy = Tmp + 1
        
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




Private Function BoundingBoxString(BBXFile As String) As String
    Const ForReading = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtStream = fso.OpenTextFile(BBXFile, ForReading, False)
    Dim TextSplit As Variant
    Do While Not txtStream.AtEndOfStream
    tmptext = txtStream.ReadLine
    TextSplit = Split(tmptext, " ")
    If TextSplit(0) = "%%HiResBoundingBox:" Then
        llx = val(TextSplit(1))
        lly = val(TextSplit(2))
        urx = val(TextSplit(3))
        ury = val(TextSplit(4))
        'compute size and offset
        sx = CStr(Round((urx - llx) / 72 * 1200))
        sy = CStr(Round((ury - lly) / 72 * 1200))
        cx = str(-llx)
        cy = str(-lly)
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
    
End Sub

Private Sub LoadSettings()
    RegPath = "Software\IguanaTex"
    checkboxTransp.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Transparent", True)
    checkboxDebug.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Debug", False)
    textboxSize.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "PointSize", "20")
    TextBox1.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexCode", "\documentclass{article}" & Chr(13) & "\usepackage{amsmath}" & Chr(13) & "\pagestyle{empty}" & Chr(13) & "\begin{document}" & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "\end{document}")
    TextBox1.SelStart = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexCodeCursor", 0)
    MultiPage1.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Multipage", 0)
    
    LaTexEngineList = Array("pdflatex", "pdflatex", "xelatex", "lualatex")
    LaTexEngineDisplayList = Array("latex (DVI->PNG)", "pdflatex (PDF->PNG)", "xelatex (PDF->PNG)", "lualatex (PDF->PNG)")
    UsePDFList = Array(False, True, True, True)
    ComboBoxLaTexEngine.List = LaTexEngineDisplayList
    ComboBoxLaTexEngine.ListIndex = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LaTeXEngineID", 0)
    
    TemplateSortedListString = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TemplateSortedList", "0")
    TemplateSortedList = UnpackStringToArray(TemplateSortedListString)
    'NumberOfTemplates = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "NumberOfTemplates", 1)
    TemplateNameSortedListString = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TemplateNameSortedList", "New Template")
    ComboBoxTemplate.List = UnpackStringToArray(TemplateNameSortedListString)
End Sub

Private Function GetTempPath() As String
    Dim res As String
    RegPath = "Software\IguanaTex"
    res = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Temp Dir", "c:\temp")
    If Right(res, 1) <> "\" Then
        res = res & "\"
    End If
    GetTempPath = res
End Function

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
    End If
    TextBoxTemplateCode.SetFocus
End Sub

Private Sub UpdateTemplateRegistry()
    ' update the list of saved templates names in the registry (will be used to initialize combo box content)
    TemplateSortedListString = PackArrayToString(TemplateSortedList)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "TemplateSortedList", REG_SZ, CStr(TemplateSortedListString)
    ' save the number of templates in registry
    'SetRegistryValue HKEY_CURRENT_USER, RegPath, "NumberOfTemplates", REG_DWORD, CLng(NumberOfTemplates)
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

Private Sub UserForm_Initialize()
    LoadSettings
    
    LatexForm.textboxSize.Visible = True
    LatexForm.Label2.Visible = True
    LatexForm.Label3.Visible = True

    FrameProcess.Visible = False
    
End Sub

Private Sub UserForm_Activate()
    'Execute macro to enable resizeability
    MakeFormResizable
    
    RegPath = "Software\IguanaTex"
    LatexForm.Height = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexFormHeight", 312)
    LatexForm.Width = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexFormWidth", 380)
    ResizeForm
    
End Sub

Sub RetrieveOldShapeInfo(oldShape As Shape, mainText As String)
    CheckBoxReset.Visible = True
    CheckBoxReset.Value = False
    Label2.Caption = "Reset size:"
    ButtonRun.Caption = "ReGenerate"
    ButtonRun.Accelerator = "G"
    
    TextBox1.Text = mainText
    CursorPosition = Len(TextBox1.Text)
    
    With oldShape.Tags
        For j = 1 To .count
            If (.name(j) = "IGUANATEXSIZE") Then
                textboxSize.Text = .Value(j)
            End If
            If (.name(j) = "TRANSPARENCY") Then
                checkboxTransp.Value = .Value(j)
            End If
            If (.name(j) = "IGUANATEXCURSOR") Then
                CursorPosition = .Value(j)
            End If
            If (.name(j) = "LATEXENGINEID") Then
                ComboBoxLaTexEngine.ListIndex = .Value(j)
            End If
        Next
    End With
    TextBox1.SelStart = CursorPosition
    textboxSize.Enabled = False
End Sub

Private Sub UserForm_Resize()
    ' Minimal size
    minLatexFormHeight = MultiPage1.Top + 18 + 50 + 4 * ButtonAbout.Height
    minLatexFormWidth = ButtonCancel.Left + ButtonCancel.Width + ButtonAbout.Width + 22
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
    TextBox1.Width = MultiPage1.Width - 10
    TextBox1.Height = MultiPage1.Height - TextBox1.Top - 20
    
    'Other elements are moved as needed
    ButtonAbout.Top = MultiPage1.Top + MultiPage1.Height + bordersize
    ButtonAbout.Left = ButtonCancel.Left + ButtonCancel.Width + bordersize
    'ButtonAbout.Left = MultiPage1.Left + TextBox1.Left + TextBox1.Width - ButtonAbout.Width
    ButtonRun.Top = ButtonAbout.Top '+ ButtonAbout.Height / 2 + 1
    ButtonCancel.Top = ButtonRun.Top
    ButtonMakeDefault.Top = ButtonAbout.Top + ButtonAbout.Height + bordersize
    ButtonMakeDefault.Left = ButtonAbout.Left
    FrameProcess.Top = ButtonMakeDefault.Top
    FrameProcess.Left = ButtonRun.Left
    LabelProcess.Width = FrameProcess.Width
    LabelProcess.Top = 4
    
    CmdButtonImportCode.Left = TextBox1.Width - CmdButtonImportCode.Width + 1
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

Private Sub MoveAnimation(oldShape As Shape, newShape As Shape)
    ' Move the animation settings of oldShape to newShape
    With ActiveWindow.Selection.SlideRange(1).TimeLine
        Dim eff As Effect
        For Each eff In .MainSequence
            If eff.Shape.name = oldShape.name Then eff.Shape = newShape
        Next
    End With
End Sub

Private Sub MatchZOrder(oldShape As Shape, newShape As Shape)
    ' Make the Z order of newShape equal to 1 higher than that of oldShape
    newShape.ZOrder msoBringToFront
    While (newShape.ZOrderPosition > oldShape.ZOrderPosition + 1)
        newShape.ZOrder msoSendBackward
    Wend
End Sub

Private Sub DeleteAnimation(oldShape As Shape)
    ' Delete the animation settings of oldShape
    With ActiveWindow.Selection.SlideRange(1).TimeLine
        For i = .MainSequence.count To 1 Step -1
            Dim eff As Effect
            Set eff = .MainSequence(i)
            If eff.Shape.name = oldShape.name Then eff.Delete
        Next
    End With
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
            'TextBox1.Visible = True
            TextBox1.SetFocus
            'ButtonTexPath.Visible = False
            'TextBoxFile.Visible = False
            'ButtonLoadFile.Visible = False
        Case 1 ' Read from file
            'TextBox1.Visible = False
            'TextBoxFile.Visible = True
            'ButtonTexPath.Visible = True
            TextBoxFile.SetFocus
            'ButtonLoadFile.Visible = True
            ButtonLoadFile.Enabled = fs.FileExists(TextBoxFile.Text) And isTex(TextBoxFile.Text)
        Case Else ' Templates
            If TextBoxTemplateName.Text = "" Then
                TextBoxTemplateName.Text = ComboBoxTemplate.Text
            End If
            TextBoxTemplateCode.SetFocus
            
    End Select
    
End Sub

Private Function PackArrayToString(vArray As Variant) As String
    Dim strDelimiter As String
    strDelimiter = "|"
    PackArrayToString = Join(vArray, strDelimiter)
End Function

Private Function UnpackStringToArray(str As String) As Variant
    Dim strDelimiter As String
    strDelimiter = "|"
    UnpackStringToArray = Split(str, strDelimiter, , vbTextCompare)
End Function
