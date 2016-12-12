VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LatexForm 
   Caption         =   "IguanaTex"
   ClientHeight    =   6167
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   8540.001
   OleObjectBlob   =   "LatexForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LatexForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim theAppEventHandler As New AppEventHandler

Sub InitializeApp()
    ' This enables us to capture application events, such as double-clicking on
    ' an IguanaTex image
    Set theAppEventHandler.App = Application
    
    AddMenuItem "New Latex e&quation...", "NewLatexEquation", 18 '226
    AddMenuItem "Edit Latex equation...", "EditLatexEquation", 37
    AddMenuItem "Set Temp folder...", "LoadSetTempForm", 548
    
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
    RemoveMenuItem "Set Temp folder..."

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


Private Sub ButtonRun_Click()
    Dim TempPath As String
    TempPath = GetTempPath()
    FilePrefix = "addin_tmp"
    
    Dim debugMode As Boolean
    If checkboxDebug.Value Then
        debugMode = True
    Else
        debugMode = False
    End If
    
    ' Write latex to a temp file
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile(TempPath + FilePrefix + ".tex", True)
    f.Write TextBox1.Text
    f.Close
    
    ' Run latex
    RetVal& = Execute("latex -interaction=batchmode """ + FilePrefix + ".tex""", TempPath, debugMode)
    
    If (RetVal& <> 0 Or Not fs.FileExists(TempPath & FilePrefix & ".dvi")) Then
        ' Error in Latex code
        ' Read log file and show it to the user
        Dim LogFile As Object
        Set LogFile = fs.OpenTextFile(TempPath + FilePrefix + ".log", ForReading)
        LogFileViewer.TextBox1.Text = LogFile.ReadAll
        LogFile.Close
        LogFileViewer.TextBox1.ScrollBars = fmScrollBarsBoth
        LogFileViewer.Show 1
        Exit Sub
    End If
    
    DviPngSwitches = "-q -D 1200 -T tight"  ' monitor is 96 dpi; add factor of 10 to allow scaling
    If checkboxTransp.Value = True Then
        DviPngSwitches = DviPngSwitches & " -bg Transparent"
    End If
    
    Execute "dvipng " & DviPngSwitches & " -o """ & FilePrefix & ".png"" """ & FilePrefix & ".dvi""", TempPath, debugMode
    
    FinalFilename = FilePrefix & ".png"
    
    ' Latex run successful.
    ' If we are in Modify mode, store parameters of old image
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim oldShape As Shape
    IsInGroup = False
    If ButtonRun.Caption = "Modify" Then
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
    ' Add tags storing the original height and width, used next time to keep resizing ratio.
    newShape.ScaleHeight 1#, msoTrue
    newShape.ScaleWidth 1#, msoTrue
    newShape.Tags.Add "ORIGINALHEIGHT", newShape.Height
    newShape.Tags.Add "ORIGINALWIDTH", newShape.Width
    
    ' Scale it
    If ButtonRun.Caption <> "Modify" Then
        PointSize = val(textboxSize.Text)
        ScaleFactor = PointSize / 100
        newShape.ScaleHeight ScaleFactor, msoTrue
        newShape.ScaleWidth ScaleFactor, msoTrue
    Else
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
    
    ' Add tags
    newShape.Tags.Add "LATEXADDIN", TextBox1.Text
    newShape.Tags.Add "IguanaTexSize", val(textboxSize.Text)
    
    ' Copy animation settings and formatting from old image, then delete it
    If ButtonRun.Caption = "Modify" Then
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
        ' Get back to original selection state
        newShape.Select
    End If
    
    ' Delete temp files
    fs.DeleteFile TempPath + FilePrefix + "*.*"
    
    Unload LatexForm
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
         ' all other elements with need to be handled either through their group
         ' name if in a group, or their name if not
        TagGroupHierarchy = 1
        For Each n In arr
            ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "LAYER", TagGroupHierarchy
            ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "SELECTIONNAME", n
        Next
    End If


End Function




Private Function FindBoundingBoxString(PSFile As Object, OutputDPI As Integer) As String
    Dim s As String
    Do
        s = PSFile.ReadLine
        If Left(s, 15) = "%%BoundingBox: " Then
            sa = Split(Mid(s, 16))
            x1 = val(sa(0))
            y1 = val(sa(1))
            x2 = val(sa(2))
            y2 = val(sa(3))
            w = Round((x2 - x1) * (OutputDPI / 72))
            h = Round((y2 - y1) * (OutputDPI / 72))
            FindBoundingBoxString = "-g" & CStr(w) & "x" & CStr(h) & " -c " & Str(-x1) & " " & Str(-y1) & " translate -q"
            Exit Function
        End If
    Loop Until (PSFile.AtEndOfStream)
End Function

Private Sub SaveSettings()
    Dim RegPath As String
    RegPath = "Software\IguanaTex"
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Transparent", REG_DWORD, BoolToInt(checkboxTransp.Value)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "Debug", REG_DWORD, BoolToInt(checkboxDebug.Value)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "PointSize", REG_DWORD, CLng(val(textboxSize.Text))
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LatexCode", REG_SZ, CStr(TextBox1.Text)
End Sub

Private Sub LoadSettings()
    RegPath = "Software\IguanaTex"
    checkboxTransp.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Transparent", True)
    checkboxDebug.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "Debug", False)
    textboxSize.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "PointSize", "20")
    TextBox1.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LatexCode", "\documentclass{article}" & Chr(13) & "\usepackage{amsmath}" & Chr(13) & "\pagestyle{empty}" & Chr(13) & "\begin{document}" & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "\end{document}")
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

Private Sub CommandButton1_Click()
    AboutBox.Show 1
End Sub

Private Sub CommandButton2_Click()
    SaveSettings
End Sub

Private Sub UserForm_Initialize()
    LoadSettings
    'This is only to make sure that the form aligns everything, this way there isn't a slight jump when the user first resizes the window
    TextBox1.Height = LatexForm.Height - CommandButton1.Height * 5
    TextBox1.Width = LatexForm.Width - 25
    
    ButtonRun.Top = LatexForm.Height - ButtonRun.Height * 3
    ButtonCancel.Top = LatexForm.Height - ButtonCancel.Height * 3
    CommandButton1.Top = LatexForm.Height - CommandButton1.Height * 4
    CommandButton2.Top = LatexForm.Height - CommandButton2.Height * 3
    
    checkboxDebug.Top = LatexForm.Height - checkboxDebug.Height * 3
    checkboxTransp.Top = LatexForm.Height - checkboxTransp.Height * 4
    checkboxTransp.Top = LatexForm.Height - checkboxTransp.Height * 4
    Label2.Top = LatexForm.Height - Label2.Height * 7
    textboxSize.Top = LatexForm.Height - Label2.Height * 7
    Label3.Top = LatexForm.Height - Label2.Height * 7
    
    
End Sub

Private Sub UserForm_Activate()
  'Execute macro to enable resizeability
  MakeFormResizable
End Sub

Private Sub UserForm_Resize()
    'Make sure that the size is not zero!
    If LatexForm.Height - CommandButton1.Height * 5 > 0 Then
        TextBox1.Height = LatexForm.Height - CommandButton1.Height * 5
        TextBox1.Width = LatexForm.Width - 25
    End If
    
    'Other elements are moved as needed
    ButtonRun.Top = LatexForm.Height - ButtonRun.Height * 3
    ButtonCancel.Top = LatexForm.Height - ButtonCancel.Height * 3
    CommandButton1.Top = LatexForm.Height - CommandButton1.Height * 4
    CommandButton2.Top = LatexForm.Height - CommandButton2.Height * 3
    
    checkboxDebug.Top = LatexForm.Height - checkboxDebug.Height * 3
    checkboxTransp.Top = LatexForm.Height - checkboxTransp.Height * 4
    checkboxTransp.Top = LatexForm.Height - checkboxTransp.Height * 4
    Label2.Top = LatexForm.Height - Label2.Height * 7
    textboxSize.Top = LatexForm.Height - Label2.Height * 7
    Label3.Top = LatexForm.Height - Label2.Height * 7
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
