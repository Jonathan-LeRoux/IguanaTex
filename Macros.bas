Attribute VB_Name = "Macros"
Option Explicit

Public RegenerateContinue As Boolean

Sub NewLatexEquation()
    Load LatexForm
    
    Dim osld As Slide
    On Error Resume Next
    Set osld = ActiveWindow.Selection.SlideRange(1)
    If Err <> 0 Then
        MsgBox "Please select a slide on which to generate the LaTeX display."
        Exit Sub
    End If
    
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    If (Sel.Type = ppSelectionText) Then
        LatexForm.textboxSize.Text = Sel.TextRange.Font.Size
    End If
    
    If IsEmpty(LatexForm.textboxSize.Text) Then
        LatexForm.textboxSize.Text = "20"
    End If
    LatexForm.CheckBoxReset.Visible = False
    LatexForm.Label2.Caption = "Set size:"
       
    LatexForm.ButtonRun.Caption = "Generate"
    LatexForm.ButtonRun.Accelerator = "G"

    LatexForm.textboxSize.Enabled = True
    
    ShowLatexForm
End Sub

Private Sub ShowLatexForm()
    Dim UseExternalEditor As Boolean
    UseExternalEditor = GetITSetting("UseExternalEditor", False)
    If UseExternalEditor Then
        Dim TempFolder As String
        Dim LatexCode As String
        TempFolder = LatexForm.TextBoxTempFolder.Text
        LatexCode = LatexForm.TextWindow1.Text
        Load ExternalEditorForm
        ExternalEditorForm.LaunchExternalEditor TempFolder, LatexCode
'        LatexForm.Show vbModeless
'        LatexForm.CmdButtonExternalEditor_Click
    Else
        LatexForm.Show vbModal
    End If
End Sub

Sub EditLatexEquation()
    ' Check if the user currently has a single Latex equation selected.
    ' If so, display the dialog box. If not, dislpay an error message.
    ' Called when the user clicks the "Edit Latex Equation" menu item.
    
    If Not TryEditLatexEquation() Then
        MsgBox "You must select a single IguanaTex display to modify it."
    End If
End Sub

Function TryEditLatexEquation() As Boolean
    ' Analyze the type of selected object to determine if it can be edited
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim oldshape As Shape
                                                        
    If Sel.Type = ppSelectionShapes Then
        ' First make sure we don't have any shapes with duplicate names on this slide
        DeDuplicateShapeNamesInSlide ActiveWindow.View.Slide.SlideIndex
        If Sel.ShapeRange.count = 1 Then ' if not 1, then multiple objects are selected
            If Sel.ShapeRange.Type = msoGroup Then
                ' Group case: either 1 object within a group, or 1 group corresponding to an EMF display
                If Sel.HasChildShapeRange = False Then ' Maybe an EMF display
                    Set oldshape = Sel.ShapeRange(1)
                    TryEditLatexEquation = TryProcessShape(oldshape)
                    Exit Function
                ElseIf Sel.ChildShapeRange.count = 1 Then
                    ' 1 object inside a group
                    Set oldshape = Sel.ChildShapeRange(1)
                    TryEditLatexEquation = TryProcessShape(oldshape)
                    Exit Function
                End If
            Else
                ' Non-group case: only a single object can be selected
                Set oldshape = Sel.ShapeRange(1)
                If oldshape.Tags.item("EMFchild") <> vbNullString Then
                    TryEditLatexEquation = False ' we should not have an EMF child object by itself
                Else
                    TryEditLatexEquation = TryProcessShape(oldshape)
                End If
                Exit Function
            End If
        End If
    End If
    
    TryEditLatexEquation = False
End Function

Private Function TryProcessShape(oldshape As Shape) As Boolean
    Dim LatexText As String
    Dim j As Long
    
    TryProcessShape = False
    With oldshape.Tags
        If .item("LATEXADDIN") <> vbNullString Then ' we're dealing with an IguanaTex display
            For j = 1 To .count
                Debug.Print .Name(j) & vbTab & .value(j)
            Next j
            Load LatexForm
            
            LatexText = .item("LATEXADDIN")
            LatexForm.RetrieveOldShapeInfo oldshape, LatexText
            ShowLatexForm
            TryProcessShape = True
            Exit Function
        ElseIf .item("SOURCE") <> vbNullString Then ' we're dealing with a Texpoint display
            For j = 1 To .count
                Debug.Print .Name(j) & vbTab & .value(j)
            Next j
            Load LatexForm
            
            LatexText = GetLatexTextFromTexPointShape(oldshape)
            LatexForm.RetrieveOldShapeInfo oldshape, LatexText
            ShowLatexForm
            TryProcessShape = True
            Exit Function
        End If
    End With
    If TryProcessShape = False Then
        'maybe a LatexIt display
        LatexText = GetLatexTextFromLatexItShape(oldshape)
        If LatexText <> vbNullString Then
            ' LatexIt display
            oldshape.Tags.Add "BitmapVector", 0
            LatexForm.RetrieveOldShapeInfo oldshape, LatexText
            ShowLatexForm
            TryProcessShape = True
            Exit Function
        End If
    End If
End Function


' Make sure there aren't multiple shapes with the same name prior to processing
Private Sub DeDuplicateShapeNamesInSlide(SlideIndex As Integer)
    Dim vSh As Shape
    Dim vSl As Slide
    Set vSl = ActivePresentation.Slides(SlideIndex)
    
    Dim NameList() As String
    
    Dim dict As New Dictionary
    Dim n As Variant
    Dim Key As String
    For Each vSh In vSl.Shapes
        If vSh.Type = msoGroup Then
            NameList = CollectGroupedItemList(vSh, True)
        Else
            ReDim NameList(0 To 0) As String
            NameList(0) = vSh.Name
        End If
        For n = LBound(NameList) To UBound(NameList)
            Key = NameList(n)
            If Not dict.Exists(Key) Then
                dict.item(Key) = 1
            Else
                dict.item(Key) = dict.item(Key) + 1
            End If
        Next n
    Next vSh
    
    For Each vSh In vSl.Shapes
        Set dict = RenameDuplicateShapes(vSh, dict)
    Next vSh
    
    
    Set dict = Nothing
End Sub

Private Function RenameDuplicateShapes(vSh As Shape, dict As Dictionary) As Dictionary
    If vSh.Type = msoGroup Then
        Dim n As Long
        For n = 1 To vSh.GroupItems.count
            Set dict = RenameDuplicateShapes(vSh.GroupItems(n), dict)
        Next
    Else
        Dim K As String
        Dim shpCount As Long
        K = vSh.Name
        If dict.item(K) > 1 Then
            shpCount = 1
            Do While dict.Exists(K & " " & shpCount)
                shpCount = shpCount + 1
            Loop
            vSh.Name = K & " " & shpCount
            dict.Add K & " " & shpCount, 1
        End If
    End If
    Set RenameDuplicateShapes = dict
End Function


Public Sub RegenerateSelectedDisplays(Sel As Selection)
    Dim vSh As Shape
    Dim vSl As Slide
    Dim SlideIndex As Integer
    Dim DisplayCount As Long

    RegenerateContinue = True
    
    Select Case Sel.Type
        Case ppSelectionShapes
            ' Regenerate 1 or more shapes on a single slide
            SlideIndex = ActiveWindow.View.Slide.SlideIndex
            DeDuplicateShapeNamesInSlide SlideIndex
            DisplayCount = CountDisplaysInSelection(Sel)
            If DisplayCount > 0 Then
                RegenerateForm.LabelSlideNumber.Caption = 1
                RegenerateForm.LabelTotalSlideNumber.Caption = 1
                RegenerateForm.LabelShapeNumber.Caption = 0
                RegenerateForm.LabelTotalShapeNumberOnSlide.Caption = DisplayCount
                RegenerateForm.Show False
                If Sel.HasChildShapeRange Then ' displays within a group
                    For Each vSh In Sel.ChildShapeRange
                        RegenerateOneDisplay vSh
                    Next vSh
                Else
                    For Each vSh In Sel.ShapeRange
                        If vSh.Type = msoGroup And Not IsShapeDisplay(vSh) Then ' grouped displays
                            RegenerateGroupedDisplays vSh, SlideIndex
                        Else ' single display
                            RegenerateOneDisplay vSh
                        End If
                    Next vSh
                End If
            Else
                MsgBox "No displays to be regenerated."
            End If
        Case ppSelectionSlides
            ' Regenerate all shapes on 1 or more slides
            RegenerateForm.LabelSlideNumber.Caption = 0
            RegenerateForm.LabelTotalSlideNumber.Caption = Sel.SlideRange.count
            RegenerateForm.LabelShapeNumber.Caption = 0
            RegenerateForm.LabelTotalShapeNumberOnSlide.Caption = 0
            RegenerateForm.Show False
            For Each vSl In Sel.SlideRange
                RegenerateForm.LabelSlideNumber.Caption = RegenerateForm.LabelSlideNumber.Caption + 1
                DisplayCount = CountDisplaysInSlide(vSl)
                RegenerateForm.LabelTotalShapeNumberOnSlide.Caption = DisplayCount
                DoEvents
                If DisplayCount > 0 Then
                    RegenerateDisplaysOnSlide vSl
                End If
            Next vSl
        Case Else
            MsgBox "You need to select a set of shapes or slides."
    End Select
    
    With RegenerateForm
        .Hide
        .LabelShapeNumber.Caption = 0
        .LabelSlideNumber.Caption = 0
        .LabelTotalSlideNumber.Caption = 0
        .LabelTotalShapeNumberOnSlide.Caption = 0
    End With
    Unload RegenerateForm
End Sub

Sub RegenerateDisplaysOnSlide(vSl As Slide)
    vSl.Select
    DeDuplicateShapeNamesInSlide vSl.SlideIndex
    Dim vSh As Shape
    For Each vSh In vSl.Shapes
        If vSh.Type = msoGroup And Not IsShapeDisplay(vSh) Then
            RegenerateGroupedDisplays vSh, vSl.SlideIndex
        Else
            RegenerateOneDisplay vSh
        End If
    Next vSh
End Sub

Sub RegenerateGroupedDisplays(vGroupSh As Shape, SlideIndex As Integer)
    Dim n As Long
    Dim vSh As Shape
    
    Dim ItemToRegenerateList() As String
    
    ItemToRegenerateList = CollectGroupedItemList(vGroupSh, False)
    
    For n = LBound(ItemToRegenerateList) To UBound(ItemToRegenerateList)
        Set vSh = ActivePresentation.Slides(SlideIndex).Shapes(ItemToRegenerateList(n))
        RegenerateOneDisplay vSh
    Next

End Sub

Private Function CollectGroupedItemList(vSh As Shape, AllDisplays As Boolean) As Variant
    Dim n As Long
    Dim j As Long
    Dim prev_length As Long
    Dim added_length As Long
    Dim TmpList() As String
    Dim SubList() As String
    Dim AddToList As Boolean
    prev_length = -1
    For n = 1 To vSh.GroupItems.count
        If vSh.GroupItems(n).Type = msoGroup Then ' this case should never occur, as PPT disregards subgroups. Consider removing.
            SubList = CollectGroupedItemList(vSh.GroupItems(n), AllDisplays)
            added_length = UBound(SubList)
            ReDim Preserve TmpList(0 To prev_length + added_length) As String
            For j = prev_length + 1 To UBound(TmpList)
                TmpList(j) = SubList(j - prev_length - 1)
            Next j
            prev_length = UBound(TmpList)
        Else
            If AllDisplays Then
                AddToList = True
            ElseIf IsShapeDisplay(vSh.GroupItems(n)) Then ' Avoid this lengthy check if possible
                AddToList = True
            Else
                AddToList = False
            End If
            If AddToList Then
                ReDim Preserve TmpList(0 To prev_length + 1) As String
                TmpList(UBound(TmpList)) = vSh.GroupItems(n).Name
                prev_length = UBound(TmpList)
            End If
        End If
    Next
    CollectGroupedItemList = TmpList
End Function

Sub RegenerateOneDisplay(vSh As Shape)
    Dim LatexText As String
    Dim DoneProcessingShape As Boolean
    DoneProcessingShape = False
    If RegenerateContinue Then
    vSh.Select
    With vSh.Tags
        If .item("LATEXADDIN") <> vbNullString Then ' we're dealing with an IguanaTex display
            DoneProcessingShape = True
            RegenerateForm.LabelShapeNumber.Caption = RegenerateForm.LabelShapeNumber.Caption + 1
            DoEvents
            Load LatexForm
            
            LatexText = .item("LATEXADDIN")
            LatexForm.RetrieveOldShapeInfo vSh, LatexText

            Apply_BatchEditSettings

            LatexForm.ButtonRun_Click
            Exit Sub
        ElseIf .item("SOURCE") <> vbNullString Then ' we're dealing with a Texpoint display
            DoneProcessingShape = True
            RegenerateForm.LabelShapeNumber.Caption = RegenerateForm.LabelShapeNumber.Caption + 1
            DoEvents
            Load LatexForm
            
            LatexText = GetLatexTextFromTexPointShape(vSh)
            LatexForm.RetrieveOldShapeInfo vSh, LatexText
            
            Apply_BatchEditSettings
            
            LatexForm.ButtonRun_Click
            Exit Sub
        End If
    End With
    If DoneProcessingShape = False Then
        'maybe a LatexIt display
        LatexText = GetLatexTextFromLatexItShape(vSh)
        If LatexText <> vbNullString Then
            ' LatexIt display
            DoneProcessingShape = True
            vSh.Tags.Add "BitmapVector", 0
            LatexForm.RetrieveOldShapeInfo vSh, LatexText
            
            Apply_BatchEditSettings
            
            LatexForm.ButtonRun_Click
            Exit Sub
        End If
    End If
    Else
        Debug.Print "Pressed Cancel"
    End If
End Sub


Private Function GetLatexTextFromLatexItShape(vSh As Shape) As String
    Dim LatexText As String
    LatexText = vbNullString
    
    Dim TempPath As String
    TempPath = CleanPath(LatexForm.TextBoxTempFolder.Text)
    If Not IsPathWritable(TempPath) Then
        MsgBox "The temporary folder is not writable, so we cannot test whether this display is a LatexIt display."
    End If
    Dim FilePrefix As String
    FilePrefix = DefaultFilePrefix
    Dim TimeOutTimeString As String
    Dim TimeOutTime As Long
    TimeOutTimeString = GetITSetting("TimeOutTime", "20") ' Wait N seconds for the processes to complete
    TimeOutTime = val(NormalizeDecimalNumber(TimeOutTimeString)) * 1000
    Dim debugMode As Boolean
    debugMode = False
    Dim RetVal As Long
    Dim latexit_metadata_command As String
    latexit_metadata_command = GetITSetting("LaTeXiT", DEFAULT_LATEXIT_METADATA_COMMAND)
    Dim picPath As String
    
    #If Mac Then
        ' On Mac, we only check if the LaTeXiT metadata extractor exists if it is in the default add-in folder.
        ' Otherwise, the Mac Sandbox asks us for permission to grant access to the executable, so we don't bother.
        Dim fs As New MacFileSystemObject
        Dim ProceedWithLaTeXiT As Boolean
        If Left(latexit_metadata_command, Len(DEFAULT_ADDIN_FOLDER)) = DEFAULT_ADDIN_FOLDER Then
            ProceedWithLaTeXiT = fs.FileExists(latexit_metadata_command)
        Else
            ProceedWithLaTeXiT = True
        End If
        If ProceedWithLaTeXiT Then
            ' Shape needs to be a picture, if so we save it as PDF
            picPath = TempPath & FilePrefix & ".pdf"
            If vSh.Type = msoPicture Then
                'vSh.Export picPath, msoPictureTypePDF
                Dim NewPres As Presentation
                Set NewPres = Presentations.Add(msoFalse)
                Dim NewSlide As Slide
                Set NewSlide = NewPres.Slides.Add(index:=1, Layout:=ppLayoutBlank)
                'Dim NewShape As Shape
                Dim ClipboardString As String
                ClipboardString = Clipboard ' Backup clipboard text if any
                vSh.Copy
                NewPres.Slides(1).Shapes.Paste
                Clipboard ClipboardString ' Restore clipboard text if any
                ' This briefly displays a saving progress dialog, but I haven't found a way to disable that
                NewPres.SaveAs picPath, ppSaveAsPDF
                'Application.DisplayAlerts = True
                NewPres.Close
                Set NewPres = Nothing
            End If
        End If
    #Else
        ' no need to go through all this trouble if the user does not have the LaTeXiT metadata extractor...
        Dim fs As New FileSystemObject
        If fs.FileExists(latexit_metadata_command) Then
            ' Shape needs to be a picture, if so we save it as EMF
            picPath = vbNullString
            If vSh.Type = msoPicture Then
                ' Unfortunately, saving as EMF corrupts the EMF file and loses the LaTeXiT info
                'vSh.Export picPath, ppShapeFormatEMF
                ' So, we need to use a more complicated route...
                Dim ClipboardString As String
                ClipboardString = Clipboard ' Backup clipboard text if any
                
                Dim NewPres As Presentation
                Set NewPres = Presentations.Add(msoFalse)
                Dim NewSlide As Slide
                Set NewSlide = NewPres.Slides.Add(index:=1, Layout:=ppLayoutBlank)
                'Dim NewShape As Shape
                vSh.Copy
                NewPres.Slides(1).Shapes.Paste
                Dim FilePrefixLatexit As String
                FilePrefixLatexit = TempPath & FilePrefix & "_latexit"
                NewPres.SaveAs (FilePrefixLatexit & ".pptx")
                NewPres.Close
                Set NewPres = Nothing
                
                Clipboard ClipboardString ' Restore clipboard text if any
                
                fs.CopyFile FilePrefixLatexit & ".pptx", FilePrefixLatexit & ".zip", True
                fs.DeleteFile FilePrefixLatexit & ".pptx"
                Dim Image1EMF As String
                Image1EMF = "ppt\media\image1.emf"
                RetVal& = Execute("unzip -o " & FilePrefixLatexit & ".zip" & " " & Image1EMF _
                                    & " -d " & FilePrefixLatexit, TempPath, debugMode, TimeOutTime)
                If fs.FileExists(FilePrefixLatexit & ".zip") Then
                    fs.DeleteFile FilePrefixLatexit & ".zip"
                End If
                If fs.FileExists(FilePrefixLatexit & "\" & Image1EMF) Then
                    fs.CopyFile FilePrefixLatexit & "\" & Image1EMF, FilePrefixLatexit & ".emf"
                    picPath = FilePrefixLatexit & ".emf"
                End If
                If fs.FolderExists(FilePrefixLatexit) Then
                    fs.DeleteFolder FilePrefixLatexit
                End If
            End If
        End If
    #End If
    
    ' Run LatexIt metadata extractor
    If fs.FileExists(picPath) Then
        Dim RunCommand As String
        RunCommand = ShellEscape(latexit_metadata_command) & " " & picPath
        RetVal& = Execute(RunCommand, TempPath, debugMode, TimeOutTime)
        If (RetVal& <> 0) Then
            Dim ErrorMessage As String
            ErrorMessage = "LatexIt Metadata extraction did not run properly, please make sure it runs on the command line."
            ShowError ErrorMessage, RunCommand
        Else
            If fs.FileExists(picPath & ".tex") Then
                LatexText = ReadAll(picPath & ".tex")
                If Len(LatexText) > 16 Then
                    vSh.Tags.Add "IGUANATEXCURSOR", Len(LatexText) - 16
                End If
                fs.DeleteFile picPath & ".tex"
            End If
            If fs.FileExists(picPath) Then fs.DeleteFile picPath
        End If
    End If
    GetLatexTextFromLatexItShape = LatexText
End Function

Private Function GetLatexTextFromTexPointShape(vSh As Shape) As String
    Dim LatexText As String
    Dim SourceParts() As String
    Dim TeXSource As String
    Dim ScalingFactor As Single
    Dim IsTemplate As Boolean
    
    With vSh.Tags
        ScalingFactor = 1
        IsTemplate = False
        If .item("ORIGWIDTH") <> vbNullString Then
            ScalingFactor = ScalingFactor * vSh.Width / val(NormalizeDecimalNumber(.item("ORIGWIDTH")))
        End If
        If .item("TEXPOINT") = "template" Then
            IsTemplate = True
        End If
        vSh.Tags.Add "TEXPOINTSCALING", ScalingFactor
        
        If IsTemplate = True Then
            SourceParts = Split(.item("SOURCE"), vbTab, , vbTextCompare)
            If UBound(SourceParts) > 2 Then
                TeXSource = SourceParts(3)
            Else
                SourceParts = Split(.item("SOURCE"), "equation", , vbTextCompare)
                SourceParts = Split(SourceParts(1), "template TP", , vbTextCompare)
                TeXSource = SourceParts(0)
            End If
            LatexText = DEFAULT_LATEX_CODE_PRE & "$" & TeXSource & "$" & DEFAULT_LATEX_CODE_POST
            vSh.Tags.Add "IGUANATEXCURSOR", Len(LatexText) - 16
        Else
            LatexText = .item("SOURCE")
        End If
        Dim j As Long
        For j = 1 To .count
            Debug.Print .Name(j) & vbTab & .value(j)
        Next j
    End With

    GetLatexTextFromTexPointShape = LatexText
End Function

Private Sub Apply_BatchEditSettings()
    If BatchEditForm.CheckBoxModifyEngine.value Then
        LatexForm.ComboBoxLaTexEngine.ListIndex = BatchEditForm.ComboBoxLaTexEngine.ListIndex
    End If
    If BatchEditForm.CheckBoxModifyTempFolder.value Then
        LatexForm.TextBoxTempFolder.Text = BatchEditForm.TextBoxTempFolder.Text
    End If
    If BatchEditForm.CheckBoxModifyBitmapVector.value Then
        LatexForm.ComboBoxBitmapVector.ListIndex = BatchEditForm.ComboBoxBitmapVector.ListIndex
    End If
    If BatchEditForm.CheckBoxModifyLocalDPI.value Then
        LatexForm.TextBoxLocalDPI.Text = BatchEditForm.TextBoxLocalDPI.Text
    End If
    If BatchEditForm.CheckBoxModifySize.value Then
        LatexForm.CheckBoxReset.value = True
        LatexForm.textboxSize.Text = BatchEditForm.textboxSize.Text
    End If
    If BatchEditForm.CheckBoxForcePreserveSize.value Then
        LatexForm.CheckBoxForcePreserveSize.value = True
    End If
    If BatchEditForm.CheckBoxModifyTransparency.value Then
        LatexForm.checkboxTransp.value = BatchEditForm.checkboxTransp.value
        LatexForm.TextBoxChooseColor.Text = BatchEditForm.TextBoxChooseColor.Text
    End If
    If BatchEditForm.CheckBoxModifyResetFormat.value Then
        LatexForm.CheckBoxResetFormat.value = BatchEditForm.CheckBoxResetFormat.value
    End If
    If BatchEditForm.CheckBoxReplace.value Then
        If BatchEditForm.TextBoxFind.Text <> vbNullString Then
            LatexForm.TextWindow1.Text = Replace(LatexForm.TextWindow1.Text, BatchEditForm.TextBoxFind.Text, BatchEditForm.TextBoxReplacement.Text)
        End If
    End If
End Sub

Function IsShapeDisplay(vSh As Shape) As Boolean
    IsShapeDisplay = False
    With vSh.Tags
        If .item("LATEXADDIN") <> vbNullString Then ' we're dealing with an IguanaTex display
            IsShapeDisplay = True
        ElseIf .item("SOURCE") <> vbNullString Then ' we're dealing with a Texpoint display
            IsShapeDisplay = True
        ElseIf .item("IsLatexItDisplay") <> vbNullString Then ' see if we've already checked if LatexIt display
            If .item("IsLatexItDisplay") = "True" Then
                IsShapeDisplay = True
            Else
                IsShapeDisplay = False
            End If
        ElseIf GetLatexTextFromLatexItShape(vSh) <> vbNullString Then ' we're dealing with a LatexIt display
            IsShapeDisplay = True
            vSh.Tags.Add "IsLatexItDisplay", "True"
        Else
            vSh.Tags.Add "IsLatexItDisplay", "False"
        End If
    End With
End Function

Function CountDisplaysInShape(vSh As Shape) As Integer
    Dim DisplayCount As Long
    DisplayCount = 0
    If vSh.Type = msoGroup Then ' grouped displays
        Dim s As Shape
        For Each s In vSh.GroupItems
            DisplayCount = DisplayCount + CountDisplaysInShape(s)
        Next
    Else ' single display
        If IsShapeDisplay(vSh) Then
            DisplayCount = 1
        End If
    End If
    CountDisplaysInShape = DisplayCount
End Function

Function CountDisplaysInSelection(Sel As Selection) As Integer
    Dim vSh As Shape
    Dim DisplayCount As Long
    
    DisplayCount = 0
    If Sel.HasChildShapeRange Then ' displays within a group
        For Each vSh In Sel.ChildShapeRange
            DisplayCount = DisplayCount + CountDisplaysInShape(vSh)
        Next vSh
    Else
        For Each vSh In Sel.ShapeRange
            DisplayCount = DisplayCount + CountDisplaysInShape(vSh)
        Next vSh
    End If
    CountDisplaysInSelection = DisplayCount
End Function

Function CountDisplaysInSlide(vSl As Slide) As Integer
    Dim vSh As Shape
    Dim DisplayCount As Long
    DisplayCount = 0
    For Each vSh In vSl.Shapes
        DisplayCount = DisplayCount + CountDisplaysInShape(vSh)
    Next vSh
    CountDisplaysInSlide = DisplayCount
End Function

Private Sub Auto_Open()
    ' Runs when the add-in is loaded
    LatexForm.InitializeApp
    Load LatexForm
    Unload LatexForm
End Sub

Private Sub Auto_Close()
    LatexForm.UnInitializeApp
End Sub

Public Sub RegenerateSelection()
    Load BatchEditForm
    BatchEditForm.Show
End Sub

Public Sub ConvertToVector()
    Load BatchEditForm
    BatchEditForm.CheckBoxModifyBitmapVector.value = True
    BatchEditForm.ComboBoxBitmapVector.Enabled = True
    BatchEditForm.ComboBoxBitmapVector.ListIndex = 1
    BatchEditForm.CheckBoxModifyPreserveSize.value = True
    BatchEditForm.CheckBoxForcePreserveSize.Enabled = True
    BatchEditForm.CheckBoxForcePreserveSize.value = True
    BatchEditForm.ButtonRun_Click
End Sub

Public Sub ConvertToBitmap()
    Load BatchEditForm
    BatchEditForm.CheckBoxModifyBitmapVector.value = True
    BatchEditForm.ComboBoxBitmapVector.Enabled = True
    BatchEditForm.ComboBoxBitmapVector.ListIndex = 0
    BatchEditForm.CheckBoxModifyPreserveSize.value = True
    BatchEditForm.CheckBoxForcePreserveSize.Enabled = True
    BatchEditForm.CheckBoxForcePreserveSize.value = True
    BatchEditForm.ButtonRun_Click
End Sub

Public Sub LoadSettingsForm()
    Load SetTempForm
    SetTempForm.Show
End Sub

Public Sub InsertVectorGraphicsFile()
    Load LoadVectorGraphicsForm
    LoadVectorGraphicsForm.ButtonPath_Click
    LoadVectorGraphicsForm.Show
End Sub

Public Sub LoadDefaultFileAndGenerate()
    
    Dim osld As Slide
    On Error Resume Next
    Set osld = ActiveWindow.Selection.SlideRange(1)
    If Err <> 0 Then
        MsgBox "Please select a slide on which to generate the LaTeX display."
        Exit Sub
    End If
    Load LatexForm
    If FileExists(LatexForm.TextBoxFile.Text) And isTex(LatexForm.TextBoxFile.Text) Then
        LatexForm.TextWindow1.Text = ReadAll(LatexForm.TextBoxFile.Text)
        LatexForm.FrameProcess.Visible = True
        LatexForm.MultiPage1.value = 0
        LatexForm.MultiPage1.Visible = True
        LatexForm.Show vbModeless
        LatexForm.Repaint
        LatexForm.ButtonRun_Click
    Else
        MsgBox "You need to set an existing LaTeX file as default file path " _
               & "using ""Make Default"" in the IguanaTeX ""Read from file"" pane."
    End If
    Unload LatexForm
End Sub
