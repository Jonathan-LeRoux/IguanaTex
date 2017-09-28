Attribute VB_Name = "Macros"
'Written: August 02, 2010
'Author:  Leith Ross
'Summary: Makes the UserForm resizable by dragging one of the sides. Place a call
'         to the macro MakeFormResizable in the UserForm's Activate event.
'Source: http://www.mrexcel.com/forum/excel-questions/485489-resize-userform.html

#If VBA7 Then
 Private Declare PtrSafe Function SetLastError _
   Lib "kernel32.dll" _
     (ByVal dwErrCode As Long) _
   As Long
   
 Public Declare PtrSafe Function GetActiveWindow _
   Lib "user32.dll" () As Long

 Private Declare PtrSafe Function GetWindowLong _
   Lib "user32.dll" Alias "GetWindowLongA" _
     (ByVal hWnd As Long, _
      ByVal nIndex As Long) _
   As Long
               
 Private Declare PtrSafe Function SetWindowLong _
   Lib "user32.dll" Alias "SetWindowLongA" _
     (ByVal hWnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) _
   As Long
 
 Private Declare PtrSafe Function GetDC Lib "user32" _
    (ByVal hWnd As Long) As Long

 Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long

 Private Declare PtrSafe Function ReleaseDC Lib "user32" _
    (ByVal hWnd As Long, ByVal hDC As Long) As Long


#Else
 Private Declare Function SetLastError _
   Lib "kernel32.dll" _
     (ByVal dwErrCode As Long) _
   As Long
   
 Public Declare Function GetActiveWindow _
   Lib "user32.dll" () As Long

 Private Declare Function GetWindowLong _
   Lib "user32.dll" Alias "GetWindowLongA" _
     (ByVal hwnd As Long, _
      ByVal nIndex As Long) _
   As Long
               
 Private Declare Function SetWindowLong _
   Lib "user32.dll" Alias "SetWindowLongA" _
     (ByVal hwnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) _
   As Long
   
 Private Declare Function GetDC Lib "User32" _
    (ByVal hwnd As Long) As Long

 Private Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long

 Private Declare Function ReleaseDC Lib "User32" _
    (ByVal hwnd As Long, ByVal hDC As Long) As Long



#End If

Public RegenerateContinue As Boolean

Private Const LOGPIXELSX = 88  'Pixels/inch in X

'A point is defined as 1/72 inches
Private Const POINTS_PER_INCH As Long = 72

'The size of a pixel, in points
Public Function PointsPerPixel() As Double

 Dim hDC As Long
 Dim lDotsPerInch As Long

 hDC = GetDC(0)
 lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
 PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
 ReleaseDC 0, hDC

End Function

'The size of a pixel, in points
Public Function lDotsPerInch() As Long

 Dim hDC As Long
'  Dim lDotsPerInch As Long

 hDC = GetDC(0)
 lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
 ReleaseDC 0, hDC

End Function

Public Sub MakeFormResizable()

  Dim lStyle As Long
  Dim hWnd As Long
  Dim RetVal
  
  Const WS_THICKFRAME = &H40000
  Const GWL_STYLE As Long = (-16)
  
    hWnd = GetActiveWindow
  
    'Get the basic window style
     lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME
     
    'Set the basic window styles
     RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)
    
    'Clear any previous API error codes
     SetLastError 0
    
    'Did the style change?
     If RetVal = 0 Then MsgBox "Unable to make UserForm Resizable."
     
End Sub

Sub NewLatexEquation()
Attribute NewLatexEquation.VB_Description = "Macro created 24.5.2007 by Zvika Ben-Haim"
    Load LatexForm
    
    If IsEmpty(LatexForm.textboxSize.Text) Then
        LatexForm.textboxSize.Text = "20"
    End If
    LatexForm.CheckBoxReset.Visible = False
    LatexForm.Label2.Caption = "Set size:"
       
    LatexForm.ButtonRun.Caption = "Generate"
    LatexForm.ButtonRun.Accelerator = "G"
    LatexForm.textboxSize.Enabled = True
    LatexForm.Show
End Sub

Sub NewLatexEquationMatchSize(ByVal size)
    Load LatexForm
    
    LatexForm.textboxSize.Text = size
    
    LatexForm.CheckBoxReset.Visible = False
    LatexForm.Label2.Caption = "Set size:"
       
    LatexForm.ButtonRun.Caption = "Generate"
    LatexForm.ButtonRun.Accelerator = "G"
    LatexForm.textboxSize.Enabled = True
    LatexForm.Show
End Sub

Sub EditLatexEquation()
    ' Check if the user currently has a single Latex equation selected.
    ' If so, display the dialog box. If not, dislpay an error message.
    ' Called when the user clicks the "Edit Latex Equation" menu item.
    
    If Not TryEditLatexEquation() Then
        MsgBox "You must select a single IguanaTex++ equation to modify it."
    End If
End Sub

Function TryEditLatexEquation() As Boolean
    ' Analyze the type of selected object to determine if it can be edited
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim oldshape As Shape
    Dim LatexText As String
    Dim SourceParts() As String
    Dim TeXSource As String
                                                       
                                                        
    If Sel.Type = ppSelectionShapes Then
        ' First make sure we don't have any shapes with duplicate names on this slide
        Call DeDuplicateShapeNamesInSlide(ActiveWindow.View.Slide.SlideIndex)
        If Sel.ShapeRange.count = 1 Then ' if not 1, then multiple objects are selected
            ' Group case: either 1 object within a group, or 1 group corresponding to an EMF display
            If Sel.ShapeRange.Type = msoGroup Then
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
            ' Non-group case: only a single object can be selected
            Else
                Set oldshape = Sel.ShapeRange(1)
                If oldshape.Tags.Item("EMFchild") <> "" Then
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

Function TryProcessShape(oldshape As Shape) As Boolean
    Dim LatexText As String
    Dim SourceParts() As String
    Dim TeXSource As String
 
    TryProcessShape = False
    With oldshape.Tags
        If .Item("LATEXADDIN") <> "" Then ' we're dealing with an IguanaTex display
            For j = 1 To .count
                Debug.Print .name(j) & vbTab & .Value(j)
            Next j
            Load LatexForm
            
            Call LatexForm.RetrieveOldShapeInfo(oldshape, .Item("LATEXADDIN"))
            
            LatexForm.Show
            TryProcessShape = True
            Exit Function
        ElseIf .Item("SOURCE") <> "" Then ' we're dealing with a Texpoint display
            For j = 1 To .count
                Debug.Print .name(j) & vbTab & .Value(j)
            Next j
            ScalingFactor = 1
            IsTemplate = False
            If .Item("ORIGWIDTH") <> "" Then
                ScalingFactor = ScalingFactor * oldshape.Width / val(.Item("ORIGWIDTH"))
            End If
            If .Item("TEXPOINT") = "template" Then
                IsTemplate = True
            End If
            oldshape.Tags.Add "TEXPOINTSCALING", ScalingFactor
        
            Load LatexForm
            
            If IsTemplate = True Then
                SourceParts = Split(.Item("SOURCE"), vbTab, , vbTextCompare)
                If UBound(SourceParts) > 2 Then
                    TeXSource = SourceParts(3)
                Else
                    SourceParts = Split(.Item("SOURCE"), "equation", , vbTextCompare)
                    SourceParts = Split(SourceParts(1), "template TP", , vbTextCompare)
                    TeXSource = SourceParts(0)
                End If
                LatexText = "\documentclass{article}" & Chr(13) & "\usepackage{amsmath}" & Chr(13) & "\pagestyle{empty}" & Chr(13) & "\begin{document}" & Chr(13) & Chr(13) & "$" & TeXSource & "$" & Chr(13) & Chr(13) & "\end{document}"
                oldshape.Tags.Add "IGUANATEXCURSOR", Len(LatexText) - 16
            Else
                LatexText = .Item("SOURCE")
            End If
            Call LatexForm.RetrieveOldShapeInfo(oldshape, LatexText)
            LatexForm.Show
            TryProcessShape = True
            Exit Function
        End If
    End With
End Function


' Make sure there aren't multiple shapes with the same name prior to processing
Sub DeDuplicateShapeNamesInSlide(SlideIndex As Integer)
    Dim vSh As Shape
    Dim vSl As Slide
    Set vSl = ActivePresentation.Slides(SlideIndex)
    
    Dim NameList() As String
    
    Dim dict As New Scripting.Dictionary
    For Each vSh In vSl.Shapes
        If vSh.Type = msoGroup Then
            NameList = CollectGroupedItemList(vSh, True)
        Else
            ReDim NameList(0 To 0) As String
            NameList(0) = vSh.name
        End If
        For n = LBound(NameList) To UBound(NameList)
            Key = NameList(n)
            If Not dict.Exists(Key) Then
                dict.Item(Key) = 1
            Else
                dict.Item(Key) = dict.Item(Key) + 1
            End If
        Next n
    Next vSh
    
    For Each vSh In vSl.Shapes
        Set dict = RenameDuplicateShapes(vSh, dict)
    Next vSh
    
    
    Set dict = Nothing
End Sub

Private Function RenameDuplicateShapes(vSh As Shape, dict As Scripting.Dictionary) As Scripting.Dictionary
    If vSh.Type = msoGroup Then
        Dim n As Long
        For n = 1 To vSh.GroupItems.count
            Set dict = RenameDuplicateShapes(vSh.GroupItems(n), dict)
        Next
    Else
        K = vSh.name
        If dict.Item(K) > 1 Then
            shpCount = 1
            Do While dict.Exists(K & " " & shpCount)
                shpCount = shpCount + 1
            Loop
            vSh.name = K & " " & shpCount
            dict.Add K & " " & shpCount, 1
        End If
    End If
    Set RenameDuplicateShapes = dict
End Function


Public Sub RegenerateSelectedDisplays()
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim vSh As Shape
    Dim vSl As Slide
    Dim SlideIndex As Integer

    RegenerateContinue = True
    
    Select Case Sel.Type
        Case ppSelectionShapes
            SlideIndex = ActiveWindow.View.Slide.SlideIndex
            Call DeDuplicateShapeNamesInSlide(SlideIndex)
            DisplayCount = CountDisplaysInSelection(Sel)
            If DisplayCount > 0 Then
                RegenerateForm.LabelSlideNumber.Caption = 1
                RegenerateForm.LabelTotalSlideNumber.Caption = 1
                RegenerateForm.LabelShapeNumber.Caption = 0
                RegenerateForm.LabelTotalShapeNumberOnSlide.Caption = DisplayCount
                RegenerateForm.Show False
                If Sel.HasChildShapeRange Then ' displays within a group
                    For Each vSh In Sel.ChildShapeRange
                        Call RegenerateOneDisplay(vSh)
                    Next vSh
                Else
                    For Each vSh In Sel.ShapeRange
                        If vSh.Type = msoGroup And Not IsShapeDisplay(vSh) Then ' grouped displays
                            Call RegenerateGroupedDisplays(vSh, SlideIndex)
                        Else ' single display
                            Call RegenerateOneDisplay(vSh)
                        End If
                    Next vSh
                End If
            Else
                MsgBox "No displays to be regenerated."
            End If
        Case ppSelectionSlides
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
                    Call RegenerateDisplaysOnSlide(vSl)
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
    Call DeDuplicateShapeNamesInSlide(vSl.SlideIndex)
    Dim vSh As Shape
    For Each vSh In vSl.Shapes
        If vSh.Type = msoGroup And Not IsShapeDisplay(vSh) Then
            Call RegenerateGroupedDisplays(vSh, vSl.SlideIndex)
        Else
            Call RegenerateOneDisplay(vSh)
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
        Call RegenerateOneDisplay(vSh)
    Next

End Sub

Private Function CollectGroupedItemList(vSh As Shape, AllDisplays As Boolean) As Variant
    Dim n As Long
    Dim i As Long
    Dim prev_length As Long
    Dim added_length As Long
    Dim TmpList() As String
    Dim SubList() As String
    prev_length = -1
    For n = 1 To vSh.GroupItems.count
'        If n = 1 Then
'            prev_length = -1
'        Else
'            prev_length = UBound(TmpList)
'        End If
        If vSh.GroupItems(n).Type = msoGroup Then ' this case should never occur, as PPT disregards subgroups. Consider removing.
            SubList = CollectGroupedItemList(vSh.GroupItems(n), AllDisplays)
            added_length = UBound(SubList)
            ReDim Preserve TmpList(0 To prev_length + added_length) As String
            For j = prev_length + 1 To UBound(TmpList)
                TmpList(j) = SubList(j - prev_length - 1)
            Next j
        Else
            If AllDisplays Or IsShapeDisplay(vSh.GroupItems(n)) Then
            ReDim Preserve TmpList(0 To prev_length + 1) As String
            TmpList(UBound(TmpList)) = vSh.GroupItems(n).name
            End If
        End If
        prev_length = UBound(TmpList)
    Next
    CollectGroupedItemList = TmpList
End Function

Sub RegenerateOneDisplay(vSh As Shape)
    If RegenerateContinue Then
    vSh.Select
    With vSh.Tags
        If .Item("LATEXADDIN") <> "" Then ' we're dealing with an IguanaTex display
            RegenerateForm.LabelShapeNumber.Caption = RegenerateForm.LabelShapeNumber.Caption + 1
            DoEvents
            Load LatexForm
            
            Call LatexForm.RetrieveOldShapeInfo(vSh, .Item("LATEXADDIN"))

            Apply_BatchEditSettings

            Call LatexForm.ButtonRun_Click
            Exit Sub
        ElseIf .Item("SOURCE") <> "" Then ' we're dealing with a Texpoint display
            RegenerateForm.LabelShapeNumber.Caption = RegenerateForm.LabelShapeNumber.Caption + 1
            DoEvents
            IsTemplate = False
            If .Item("ORIGWIDTH") <> "" Then
                vSh.Tags.Add "TEXPOINTSCALING", vSh.Width / val(.Item("ORIGWIDTH"))
            End If
            If .Item("TEXPOINT") = "template" Then
                IsTemplate = True
            End If
            Load LatexForm
            
            Dim LatexText As String
            If IsTemplate = True Then
                Dim TeXSource As String
                Dim SourceParts() As String
                SourceParts = Split(.Item("SOURCE"), vbTab, , vbTextCompare)
                If UBound(SourceParts) > 2 Then
                    TeXSource = SourceParts(3)
                Else
                    SourceParts = Split(.Item("SOURCE"), "equation", , vbTextCompare)
                    SourceParts = Split(SourceParts(1), "template TP", , vbTextCompare)
                    TeXSource = SourceParts(0)
                End If
                LatexText = "\documentclass{article}" & Chr(13) & "\usepackage{amsmath}" & Chr(13) & "\pagestyle{empty}" & Chr(13) & "\begin{document}" & Chr(13) & Chr(13) & "$" & TeXSource & "$" & Chr(13) & Chr(13) & "\end{document}"
                vSh.Tags.Add "IGUANATEXCURSOR", Len(LatexText) - 16
            Else
                LatexText = .Item("SOURCE")
            End If
            Call LatexForm.RetrieveOldShapeInfo(vSh, LatexText)
            
            Apply_BatchEditSettings
            
            Call LatexForm.ButtonRun_Click
            Exit Sub
        End If
    End With
    Else
        Debug.Print "Pressed Cancel"
    End If
End Sub

Sub Apply_BatchEditSettings()
    If BatchEditForm.CheckBoxModifyEngine.Value Then
        LatexForm.ComboBoxLaTexEngine.ListIndex = BatchEditForm.ComboBoxLaTexEngine.ListIndex
    End If
    If BatchEditForm.CheckBoxModifyTempFolder.Value Then
        LatexForm.TextBoxTempFolder.Text = BatchEditForm.TextBoxTempFolder.Text
    End If
    If BatchEditForm.CheckBoxModifyBitmapVector.Value Then
        LatexForm.ComboBoxBitmapVector.ListIndex = BatchEditForm.ComboBoxBitmapVector.ListIndex
    End If
    If BatchEditForm.CheckBoxModifyLocalDPI.Value Then
        LatexForm.TextBoxLocalDPI.Text = BatchEditForm.TextBoxLocalDPI.Text
    End If
    If BatchEditForm.CheckBoxModifySize.Value Then
        LatexForm.CheckBoxReset.Value = True
        LatexForm.textboxSize.Text = BatchEditForm.textboxSize.Text
    End If
    If BatchEditForm.CheckBoxModifyTransparency.Value Then
        LatexForm.checkboxTransp.Value = BatchEditForm.checkboxTransp.Value
    End If
    If BatchEditForm.CheckBoxModifyResetFormat.Value Then
        LatexForm.CheckBoxResetFormat.Value = BatchEditForm.CheckBoxResetFormat.Value
    End If
    If BatchEditForm.CheckBoxReplace.Value Then
        If BatchEditForm.TextBoxFind.Text <> "" Then
            LatexForm.TextBox1.Text = Replace(LatexForm.TextBox1.Text, BatchEditForm.TextBoxFind.Text, BatchEditForm.TextBoxReplacement.Text)
        End If
    End If
End Sub

Function IsShapeDisplay(vSh As Shape) As Boolean
    IsShapeDisplay = False
    With vSh.Tags
        If .Item("LATEXADDIN") <> "" Then ' we're dealing with an IguanaTex display
            IsShapeDisplay = True
        ElseIf .Item("SOURCE") <> "" Then ' we're dealing with a Texpoint display
            IsShapeDisplay = True
        End If
    End With
End Function

Function CountDisplaysInShape(vSh As Shape) As Integer
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
    DisplayCount = 0
    For Each vSh In vSl.Shapes
        DisplayCount = DisplayCount + CountDisplaysInShape(vSh)
    Next vSh
    CountDisplaysInSlide = DisplayCount
End Function

Sub Auto_Open()
    ' Runs when the add-in is loaded
    LatexForm.InitializeApp
    Load LatexForm
    Unload LatexForm
End Sub

Sub Auto_Close()
    LatexForm.UnInitializeApp
End Sub

Sub LoadSetTempForm()
    Load SetTempForm
    SetTempForm.Show
End Sub

Public Sub RibbonNewLatexEquation(ByVal control)
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    If (Sel.Type = ppSelectionText) Then
        NewLatexEquationMatchSize Sel.TextRange.Font.size
    Else
        NewLatexEquation
    End If
End Sub

Public Sub RibbonEditLatexEquation(ByVal control)
    EditLatexEquation
End Sub

Public Sub RibbonSetTempFolder(ByVal control)
    LoadSetTempForm
End Sub

Public Sub RibbonRegenerateSelectedDisplays(ByVal control)
    Load BatchEditForm
    BatchEditForm.Show
End Sub

Public Sub RibbonConvertToEMF(ByVal control)
    Load BatchEditForm
    BatchEditForm.CheckBoxModifyBitmapVector.Value = True
    BatchEditForm.ComboBoxBitmapVector.Enabled = True
    BatchEditForm.ComboBoxBitmapVector.ListIndex = 1
    Call BatchEditForm.ButtonRun_Click
End Sub

Public Sub RibbonConvertToPNG(ByVal control)
    Load BatchEditForm
    BatchEditForm.CheckBoxModifyBitmapVector.Value = True
    BatchEditForm.ComboBoxBitmapVector.Enabled = True
    BatchEditForm.ComboBoxBitmapVector.ListIndex = 0
    Call BatchEditForm.ButtonRun_Click
End Sub


' Same Subs, but to be called from add-in menu in older versions of PowerPoint
Public Sub RegenerateSelectedDisplaysNoChange()
    Load BatchEditForm
    BatchEditForm.Show
End Sub

Public Sub ConvertToEMF()
    Load BatchEditForm
    BatchEditForm.CheckBoxModifyBitmapVector.Value = True
    BatchEditForm.ComboBoxBitmapVector.Enabled = True
    BatchEditForm.ComboBoxBitmapVector.ListIndex = 1
    Call BatchEditForm.ButtonRun_Click
End Sub

Public Sub ConvertToPNG()
    Load BatchEditForm
    BatchEditForm.CheckBoxModifyBitmapVector.Value = True
    BatchEditForm.ComboBoxBitmapVector.Enabled = True
    BatchEditForm.ComboBoxBitmapVector.ListIndex = 0
    Call BatchEditForm.ButtonRun_Click
End Sub

Public Sub RibbonInsertVectorGraphicsFile()
    Load LoadVectorGraphicsForm
    Call LoadVectorGraphicsForm.ButtonPath_Click
    LoadVectorGraphicsForm.Show
End Sub


Public Function GetFilePrefix() As String
    GetFilePrefix = "IguanaTex_tmp"
End Function

Public Function GetTempPath() As String
    Dim res As String
    RegPath = "Software\IguanaTex"
    res = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Temp Dir", "c:\temp\")
    If Right(res, 1) <> "\" Then
        res = res & "\"
    End If
    GetTempPath = res
End Function


Public Function GetEditorPath() As String
    Dim res As String
    RegPath = "Software\IguanaTex"
    res = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Editor", "texstudio.exe")
    GetEditorPath = res
End Function

