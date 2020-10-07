VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadVectorGraphicsForm 
   Caption         =   "Load Vector Graphics File"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6888
   OleObjectBlob   =   "LoadVectorGraphicsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadVectorGraphicsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastPath As String

Private Sub CommandButtonSave_Click()
    Dim RegPath As String
    RegPath = "Software\IguanaTex"
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LoadVectorFileConvertLines", REG_DWORD, BoolToInt(CheckBoxConvertLines.Value)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LoadVectorFileScaling", REG_SZ, textboxScalor.Text
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LoadVectorFileCalibrationX", REG_SZ, TextBoxCalibrationX.Text
    SetRegistryValue HKEY_CURRENT_USER, RegPath, "LoadVectorFileCalibrationY", REG_SZ, TextBoxCalibrationY.Text
End Sub

Private Function BoolToInt(val) As Long
    If val Then
        BoolToInt = 1&
    Else
        BoolToInt = 0&
    End If
End Function


Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    Dim RegPath As String
    RegPath = "Software\IguanaTex"
    textboxScalor.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LoadVectorFileScaling", "1")
    CheckBoxConvertLines.Value = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LoadVectorFileConvertLines", False)
    TextBoxCalibrationX.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LoadVectorFileCalibrationX", "1")
    TextBoxCalibrationY.Text = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "LoadVectorFileCalibrationY", "1")
End Sub

Private Sub ButtonCancel_Click()
    Unload LoadVectorGraphicsForm
End Sub

Private Function isEpsEmf(file As String)
    Ext = LCase(Right$(file, 3))
    If Ext = "eps" Or Ext = "emf" Or Ext = "pdf" Or Ext = ".ps" Then
        isEpsEmf = True
    Else
        isEpsEmf = False
    End If
End Function

Sub ButtonPath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Vector graphics files", "*.pdf;*.ps;*.eps;*.emf", 1
    fd.ButtonName = "&Select file"
    
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            TextBoxFile.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If

    Set fd = Nothing
    TextBoxFile.SetFocus
    
End Sub

Private Sub TextBoxFile_Change()
    Set fs = CreateObject("Scripting.FileSystemObject")
    ButtonLoadFile.Enabled = fs.FileExists(TextBoxFile.Text) And isEpsEmf(TextBoxFile.Text)
End Sub

Private Sub ButtonLoadFile_Click()
    Call InsertVectorGraphicsFile
    Unload LoadVectorGraphicsForm
End Sub


Public Sub InsertVectorGraphicsFile()
    Dim PosX As Single, PosY As Single, ScalingX As Single, ScalingY As Single
    PosX = 200
    PosY = 200
    Dim newShape As Shape
    Dim TimeOutTimeString As String
    Dim TimeOutTime As Long
    RegPath = "Software\IguanaTex"
    TimeOutTimeString = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TimeOutTime", "20") ' Wait 20 seconds for the processes to complete
    TimeOutTime = val(TimeOutTimeString) * 1000
    Dim debugMode As Boolean
    debugMode = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim StartFolder As String
    If ActivePresentation.path <> "" Then
        StartFolder = ActivePresentation.path
    Else
        StartFolder = "C:\"
    End If
    
    ' Get the path and extension of the file to be inserted
    Dim path As String, Ext As String, pdfPath As String, psPath As String
    path = TextBoxFile.Text
    Ext = LCase(Right$(path, 3))
    
    Dim DeleteTmpPDF As Boolean
    DeleteTmpPDF = False
    
    ' If .ps file, convert to .pdf first, using ps2pdf
    If Ext = ".ps" Then
        psPath = path + "_tmp.ps"
        pdfPath = path + "_tmp.pdf"
        fs.CopyFile path, psPath
        If fs.FileExists(pdfPath) Then fs.DeleteFile pdfPath
        RetVal& = Execute("""ps2pdf"" """ + psPath + """ " + pdfPath + """", StartFolder, debugMode, TimeOutTime)
        If (RetVal& <> 0 Or Not fs.FileExists(pdfPath)) Then
            MsgBox "PS to PDF conversion failed" _
            & vbNewLine & "Make sure ps2pdf.exe is installed (it comes with, e.g., Tex Live, MikTeX or Ghostscript) and can be run from anywhere via the command line"
            Exit Sub
        End If
        Ext = "pdf"
        path = pdfPath
        DeleteTmpPDF = True
    End If
    ' If .eps file, convert to .pdf first, using epspdf
    If Ext = "eps" Then
        psPath = path + "_tmp.eps"
        pdfPath = path + "_tmp.pdf"
        fs.CopyFile path, psPath
        If fs.FileExists(pdfPath) Then fs.DeleteFile pdfPath
        RetVal& = Execute("""epspdf"" """ + psPath + """ " + pdfPath + """", StartFolder, debugMode, TimeOutTime)
        If (RetVal& <> 0 Or Not fs.FileExists(pdfPath)) Then
            MsgBox " EPS to PDF conversion failed" _
            & vbNewLine & "Make sure epspdf.exe is installed (it comes with Tex Live or MikTeX) and can be run from anywhere via the command line"
            Exit Sub
        End If
        Ext = "pdf"
        path = pdfPath
        DeleteTmpPDF = True
    End If
    ' Now we're either dealing with a .pdf file or a .emf file
    
    ' If .pdf file, convert to .emf first, using pdfiumdraw, which is part of TeX2img
    If Ext = "pdf" Then
        Dim emfPath As String
        emfPath = Left$(path, Len(path) - 3) + "emf"
        Dim TmpPath As String
        TmpPath = path + "_copy.emf"
        If fs.FileExists(TmpPath) Then fs.DeleteFile TmpPath
        If fs.FileExists(emfPath) Then
            fs.CopyFile emfPath, TmpPath
            fs.DeleteFile emfPath
        End If
        tex2img_command = GetRegistryValue(HKEY_CURRENT_USER, RegPath, "TeX2img Command", "%USERPROFILE%\Downloads\TeX2img\TeX2imgc.exe")
        pdfiumdraw_command = Left$(tex2img_command, Len(tex2img_command) - Len("TeX2imgc.exe")) + "pdfiumdraw.exe"
        RetVal& = Execute("""" & pdfiumdraw_command & """ --extent=50 --emf --transparent --pages=1 """ + path + """", StartFolder, debugMode, TimeOutTime)
        If (RetVal& <> 0 Or Not fs.FileExists(emfPath)) Then
            MsgBox " PDF to EMF conversion failed" _
            & vbNewLine & "Make sure to correctly set the path to Tex2imgc.exe in Main Settings." _
            & vbNewLine & "IguanaTex uses that path to find pdfiumdraw.exe."
            Exit Sub
        End If
        Ext = "emf"
        Set newShape = AddDisplayShape(emfPath, PosX, PosY)
        If debugMode Then
            If fs.FileExists(TmpPath) Then ' Need to swap _copy.emf and the newly created file
                Dim TmpTmpPath As String
                TmpTmpPath = TmpPath + "_copy.emf"
                fs.CopyFile emfPath, TmpTmpPath
                fs.DeleteFile emfPath
                fs.CopyFile TmpPath, emfPath
                fs.DeleteFile TmpPath
                fs.CopyFile TmpTmpPath, TmpPath
                fs.DeleteFile TmpTmpPath
            End If
        Else 'Clean up
            If fs.FileExists(emfPath) Then fs.DeleteFile emfPath
            If DeleteTmpPDF Then
                If fs.FileExists(pdfPath) Then fs.DeleteFile pdfPath
            End If
            If fs.FileExists(TmpPath) Then
                fs.CopyFile TmpPath, emfPath
                fs.DeleteFile TmpPath
            End If
        End If
    Else
        Set newShape = AddDisplayShape(path, PosX, PosY)
    End If
    
    
'    If Ext = "emf" Then
'        dpi = lDotsPerInch
'        default_screen_dpi = 96
'        If dpi <> default_screen_dpi Then
'            Dim VectorScalingX As Single, VectorScalingY As Single
'            VectorScalingX = 2 * dpi / default_screen_dpi '* val(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "VectorScalingX", "1"))
'            VectorScalingY = 2 * dpi / default_screen_dpi '* val(GetRegistryValue(HKEY_CURRENT_USER, RegPath, "VectorScalingY", "1"))
'            ScalingX = textboxScalor.Value * VectorScalingX
'            ScalingY = textboxScalor.Value * VectorScalingY
'        Else
'            ScalingX = textboxScalor.Value
'            ScalingY = textboxScalor.Value
'        End If
'
'    Else
'        ScalingX = textboxScalor.Value
'        ScalingY = textboxScalor.Value
'    End If
    ScalingX = textboxScalor.Value * TextBoxCalibrationX.Value
    ScalingY = textboxScalor.Value * TextBoxCalibrationY.Value
    
    Dim ConvertLines As Boolean
    ConvertLines = CheckBoxConvertLines.Value
    Set newShape = ConvertEMF(newShape, ScalingX, ScalingY, Ext, ConvertLines)
    newShape.Select
End Sub

Private Function ConvertEMF(inSh As Shape, ScalingX As Single, ScalingY As Single, _
                            Optional FileType As String = "emf", Optional ConvertLines As Boolean = True) As Shape
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
    Dim Shr As ShapeRange
    Set Shr = inSh.Ungroup
    If FileType = "emf" Then
        Set Shr = Shr.Ungroup
        ' Clean up
        Shr.Item(1).Delete
        Shr.Item(2).Delete
        If Shr(3).GroupItems.count > 2 Then
            Set newShape = Shr(3)
        Else ' only a single freeform, so not a group
            Set newShape = Shr(3).GroupItems(2)
        End If
        Shr(3).GroupItems(1).Delete
    ElseIf FileType = "eps" Then
        Shr.GroupItems(1).Delete
        Shr.GroupItems(1).Delete
        Set newShape = Shr.Ungroup.Group
    End If
    
    
    If newShape.Type = msoGroup Then
    
        Dim arr_group() As Variant
        arr_group = GetAllShapesInGroup(newShape)
        Call FullyUngroupShape(newShape)
        Set newShape = sld.Shapes.Range(arr_group).Group
        
        Dim emf_arr() As Variant ' gather all shapes to be regrouped later on
        j_emf = 0
        Dim delete_arr() As Variant ' gather all shapes to be deleted later on
        j_delete = 0
        Dim s As Shape
        For Each s In newShape.GroupItems
            j_emf = j_emf + 1
            ReDim Preserve emf_arr(1 To j_emf)
            If s.Type = msoLine Then
                If ConvertLines And (s.Height > 0 Or s.Width > 0) Then
                    emf_arr(j_emf) = LineToFreeform(s).name
                    j_delete = j_delete + 1
                    ReDim Preserve delete_arr(1 To j_delete)
                    delete_arr(j_delete) = s.name
                Else
                    emf_arr(j_emf) = s.name
                End If
            Else
                emf_arr(j_emf) = s.name
                If s.Fill.Visible = msoTrue Then
                s.Line.Visible = msoFalse
                Else
                s.Line.Visible = msoTrue
                
                End If
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

Private Sub FullyUngroupShape(newShape As Shape)
    Dim Shr As ShapeRange
    Dim s As Shape
    If newShape.Type = msoGroup Then
        Set Shr = newShape.Ungroup
        For i = 1 To Shr.count
            Set s = Shr.Item(i)
            If s.Type = msoGroup Then
                Call FullyUngroupShape(s)
            End If
        Next
    End If
End Sub

Private Function GetAllShapesInGroup(newShape As Shape) As Variant
    Dim arr() As Variant
    Dim j As Long
    Dim s As Shape
    For Each s In newShape.GroupItems
            j = j + 1
            ReDim Preserve arr(1 To j)
            arr(j) = s.name
    Next
    GetAllShapesInGroup = arr
End Function

Private Function LineToFreeform(s As Shape) As Shape
    t = s.Line.Weight
    Dim ApplyTransform As Boolean
    ApplyTransform = True
    
    Dim bHflip As Boolean
    Dim bVflip As Boolean
    Dim nBegin As Long
    Dim nEnd As Long
    Dim aC(1 To 4, 1 To 2) As Double
    
    With s
        aC(1, 1) = .Left:           aC(1, 2) = .Top
        aC(2, 1) = .Left + .Width:  aC(2, 2) = .Top
        aC(3, 1) = .Left:           aC(3, 2) = .Top + .Height
        aC(4, 1) = .Left + .Width:  aC(4, 2) = .Top + .Height
    
        bHflip = .HorizontalFlip
        bVflip = .VerticalFlip
    End With
    
    If bHflip = bVflip Then
        If bVflip = False Then
            ' down to right -- South-East
            nBegin = 1: nEnd = 4
        Else
            ' up to left -- North-West
            nBegin = 4: nEnd = 1
        End If
    ElseIf bHflip = False Then
        ' up to right -- North-East
        nBegin = 3: nEnd = 2
    Else
        ' down to left -- South-West
        nBegin = 2: nEnd = 3
    End If
    xs = aC(nBegin, 1)
    ys = aC(nBegin, 2)
    xe = aC(nEnd, 1)
    ye = aC(nEnd, 2)
    
    ' Get unit vector in orthogonal direction
    xd = xe - xs
    yd = ye - ys
    
    s_length = Sqr(xd * xd + yd * yd)
    If s_length > 0 Then
    n_x = -yd / s_length
    n_y = xd / s_length
    Else
    n_x = 0
    n_y = 0
    End If
    
    x1 = xs + n_x * t / 2
    y1 = ys + n_y * t / 2
    x2 = xe + n_x * t / 2
    y2 = ye + n_y * t / 2
    x3 = xe - n_x * t / 2
    y3 = ye - n_y * t / 2
    x4 = xs - n_x * t / 2
    y4 = ys - n_y * t / 2
        
    'End If
    
    
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
        oSh.Fill.Visible = msoTrue
        oSh.Line.Visible = msoFalse
        oSh.Rotation = s.Rotation
        Set LineToFreeform = oSh
    Else
        Set LineToFreeform = s
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


