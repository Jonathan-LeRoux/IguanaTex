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
     (ByVal hwnd As Long, _
      ByVal nIndex As Long) _
   As Long
               
 Private Declare PtrSafe Function SetWindowLong _
   Lib "user32.dll" Alias "SetWindowLongA" _
     (ByVal hwnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) _
   As Long
 
 Private Declare PtrSafe Function GetDC Lib "User32" _
    (ByVal hwnd As Long) As Long

 Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long

 Private Declare PtrSafe Function ReleaseDC Lib "User32" _
    (ByVal hwnd As Long, ByVal hDC As Long) As Long


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
  Dim hwnd As Long
  Dim RetVal
  
  Const WS_THICKFRAME = &H40000
  Const GWL_STYLE As Long = (-16)
  
    hwnd = GetActiveWindow
  
    'Get the basic window style
     lStyle = GetWindowLong(hwnd, GWL_STYLE) Or WS_THICKFRAME
     
    'Set the basic window styles
     RetVal = SetWindowLong(hwnd, GWL_STYLE, lStyle)
    
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

Sub EditLatexEquation()
    ' Check if the user currently has a single Latex equation selected.
    ' If so, display the dialog box. If not, dislpay an error message.
    ' Called when the user clicks the "Edit Latex Equation" menu item.
    
    If Not TryEditLatexEquation() Then
        MsgBox "You must select a single IguanaTex++ equation to modify it."
    End If
End Sub

Function TryEditLatexEquation() As Boolean
    ' If the user currently has a single Latex equation selected,
    ' then open the dialog box to edit it, and return True.
    ' Otherwise, do nothing and return False.
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim oldShape As Shape
    Dim LatexText As String
    Dim SourceParts() As String
                            
                                                        
    If Sel.Type = ppSelectionShapes Then
        ' First make sure we don't have any shapes with duplicate names on this slide
        Call DeDuplicateShapeNamesInSlide(ActiveWindow.View.Slide.SlideIndex)
        ' Attempt to deal with the case of 1 object inside a group
        If Sel.ShapeRange.Type = msoGroup Then
            If Sel.ChildShapeRange.count = 1 Then
                Set oldShape = Sel.ChildShapeRange(1)
                With oldShape.Tags
                    For i = 1 To .count
                        If (.name(i) = "LATEXADDIN") Then
                            Load LatexForm
                            
                            Call LatexForm.RetrieveOldShapeInfo(oldShape, .Value(i))
                        
                            LatexForm.Show
                            TryEditLatexEquation = True
                            Exit Function
                        End If
                        If (.name(i) = "SOURCE") Then ' we're dealing with a Texpoint display
                            ScalingFactor = 1
                            IsTemplate = False
                            For j = 1 To .count
                                'Debug.Print .Name(j) & vbTab & .Value(j)
                                If (.name(j) = "ORIGWIDTH") Then
                                    ScalingFactor = ScalingFactor * oldShape.Width / .Value(j)
                                End If
                                If (.name(j) = "TEXPOINT") Then
                                    If .Value(j) = "template" Then
                                        IsTemplate = True
                                    End If
                                End If
                            Next j
                            oldShape.Tags.Add "TEXPOINTSCALING", ScalingFactor
                            Load LatexForm
                            
                            If IsTemplate = True Then
                                SourceParts = Split(.Value(i), vbTab, , vbTextCompare)
                                LatexText = "\documentclass{article}" & Chr(13) & "\usepackage{amsmath}" & Chr(13) & "\pagestyle{empty}" & Chr(13) & "\begin{document}" & Chr(13) & Chr(13) & "$" & SourceParts(3) & "$" & Chr(13) & Chr(13) & "\end{document}"
                                oldShape.Tags.Add "IGUANATEXCURSOR", Len(LatexText) - 16
                            Else
                                LatexText = .Value(i)
                            End If
                            Call LatexForm.RetrieveOldShapeInfo(oldShape, LatexText)
                            LatexForm.Show
                            TryEditLatexEquation = True
                            Exit Function
                        End If
                    Next
                End With
            End If
        ' Now the non-group case: only a single object can be selected
        ElseIf Sel.ShapeRange.count = 1 Then
            Set oldShape = Sel.ShapeRange(1)
            With oldShape.Tags
                For i = 1 To .count
                    If (.name(i) = "LATEXADDIN") Then
                        'For j = 1 To .Count
                        '    Debug.Print .Name(j) & vbTab & .Value(j)
                        'Next j
                        Load LatexForm
                        
                        Call LatexForm.RetrieveOldShapeInfo(oldShape, .Value(i))
                        
                        LatexForm.Show
                        TryEditLatexEquation = True
                        Exit Function
                    End If
                    If (.name(i) = "SOURCE") Then ' we're dealing with a Texpoint display
                        ScalingFactor = 1
                        IsTemplate = False
                        For j = 1 To .count
                            Debug.Print .name(j) & vbTab & .Value(j)
                            If (.name(j) = "ORIGWIDTH") Then
                                ScalingFactor = ScalingFactor * oldShape.Width / val(.Value(j))
                            End If
                            If (.name(j) = "TEXPOINT") Then
                                If .Value(j) = "template" Then
                                    IsTemplate = True
                                End If
                            End If
                            'If (.Name(j) = "RES") Then
                            '    ScalingFactor = ScalingFactor * 1200 / val(.Value(j))
                            'End If
                        Next j
                        oldShape.Tags.Add "TEXPOINTSCALING", ScalingFactor
                    
                        Load LatexForm
                        
                        If IsTemplate = True Then
                            SourceParts = Split(.Value(i), vbTab, , vbTextCompare)
                            LatexText = "\documentclass{article}" & Chr(13) & "\usepackage{amsmath}" & Chr(13) & "\pagestyle{empty}" & Chr(13) & "\begin{document}" & Chr(13) & Chr(13) & "$" & SourceParts(3) & "$" & Chr(13) & Chr(13) & "\end{document}"
                            oldShape.Tags.Add "IGUANATEXCURSOR", Len(LatexText) - 16
                        Else
                            LatexText = .Value(i)
                        End If
                        Call LatexForm.RetrieveOldShapeInfo(oldShape, LatexText)
                        LatexForm.Show
                        TryEditLatexEquation = True
                        Exit Function
                    End If
                Next
            End With
        End If
    End If
    
    TryEditLatexEquation = False
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
            NameList = CollectGroupedItemList(vSh)
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


Sub RegenerateSelectedDisplays()
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim vSh As Shape
    Dim vSl As Slide
    Dim SlideIndex As Integer

    Select Case Sel.Type
        Case ppSelectionShapes
            SlideIndex = ActiveWindow.View.Slide.SlideIndex
            Call DeDuplicateShapeNamesInSlide(SlideIndex)
            If Sel.HasChildShapeRange Then ' displays within a group
                For Each vSh In Sel.ChildShapeRange
                    Call RegenerateOneDisplay(vSh)
                Next vSh
            Else
                For Each vSh In Sel.ShapeRange
                    If vSh.Type = msoGroup Then ' grouped displays
                        Call RegenerateGroupedDisplays(vSh, SlideIndex)
                    Else ' single display
                        Call RegenerateOneDisplay(vSh)
                    End If
                Next vSh
            End If
        Case ppSelectionSlides
            For Each vSl In Sel.SlideRange
                Call RegenerateDisplaysOnSlide(vSl)
            Next vSl
        Case Else
            MsgBox "You need to select a set of shapes or slides."
    End Select
    
End Sub

Sub RegenerateDisplaysOnSlide(vSl As Slide)
    vSl.Select
    Call DeDuplicateShapeNamesInSlide(vSl.SlideIndex)
    Dim vSh As Shape
    For Each vSh In vSl.Shapes
        If vSh.Type = msoGroup Then
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
    
    ItemToRegenerateList = CollectGroupedItemList(vGroupSh)
    
    For n = LBound(ItemToRegenerateList) To UBound(ItemToRegenerateList)
        Set vSh = ActivePresentation.Slides(SlideIndex).Shapes(ItemToRegenerateList(n))
        Call RegenerateOneDisplay(vSh)
    Next

End Sub

Private Function CollectGroupedItemList(vSh As Shape) As Variant
    Dim n As Long
    Dim i As Long
    Dim prev_length As Long
    Dim added_length As Long
    Dim TmpList() As String
    Dim SubList() As String
    For n = 1 To vSh.GroupItems.count
        If n = 1 Then
            prev_length = -1
        Else
            prev_length = UBound(TmpList)
        End If
        If vSh.GroupItems(n).Type = msoGroup Then ' this case should never occur, as PPT disregards subgroups. Consider removing.
            SubList = CollectGroupedItemList(vSh.GroupItems(n))
            added_length = UBound(SubList)
            ReDim Preserve TmpList(0 To prev_length + added_length) As String
            For j = prev_length + 1 To UBound(TmpList)
                TmpList(j) = SubList(j - prev_ubound - 1)
            Next j
        Else
            ReDim Preserve TmpList(0 To prev_length + 1) As String
            TmpList(UBound(TmpList)) = vSh.GroupItems(n).name
        End If
    Next
    CollectGroupedItemList = TmpList
End Function

Sub RegenerateOneDisplay(vSh As Shape)
    vSh.Select
    With vSh.Tags
        For i = 1 To .count
            If (.name(i) = "LATEXADDIN") Then
                Load LatexForm
                
                Call LatexForm.RetrieveOldShapeInfo(vSh, .Value(i))
                
                Call LatexForm.ButtonRun_Click
                Exit Sub
            End If
            If (.name(i) = "SOURCE") Then ' we're dealing with a Texpoint display
                IsTemplate = False
                For j = 1 To .count
                    If (.name(j) = "ORIGWIDTH") Then
                        vSh.Tags.Add "TEXPOINTSCALING", vSh.Width / .Value(j)
                    End If
                    If (.name(j) = "TEXPOINT") Then
                        If .Value(j) = "template" Then
                            IsTemplate = True
                        End If
                    End If
                Next j
                
                Load LatexForm
                
                Dim LatexText As String
                If IsTemplate = True Then
                    Dim SourceParts() As String
                    SourceParts = Split(.Value(i), vbTab, , vbTextCompare)
                    LatexText = "\documentclass{article}" & Chr(13) & "\usepackage{amsmath}" & Chr(13) & "\pagestyle{empty}" & Chr(13) & "\begin{document}" & Chr(13) & Chr(13) & "$" & SourceParts(3) & "$" & Chr(13) & Chr(13) & "\end{document}"
                    vSh.Tags.Add "IGUANATEXCURSOR", Len(LatexText) - 16
                Else
                    LatexText = .Value(i)
                End If
                Call LatexForm.RetrieveOldShapeInfo(vSh, LatexText)
                
                Call LatexForm.ButtonRun_Click
                Exit Sub
            End If
        Next
    End With
End Sub

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
    NewLatexEquation
End Sub

Public Sub RibbonEditLatexEquation(ByVal control)
    EditLatexEquation
End Sub

Public Sub RibbonSetTempFolder(ByVal control)
    LoadSetTempForm
End Sub

Public Sub RibbonRegenerateSelectedDisplays(ByVal control)
    Call RegenerateSelectedDisplays
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
