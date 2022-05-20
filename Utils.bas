Attribute VB_Name = "Utils"
Option Explicit
Option Private Module

Public Const RegPath As String = "Software\IguanaTex" ' Registry path root
Public Const DefaultFilePrefix As String = "IguanaTex_tmp" ' Default prefix for temporary files

Private Const LOGPIXELSX = 88  'Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72 'A point is defined as 1/72 inches

#If Mac Then
Private Declare PtrSafe Function MacMakeFormResizable _
 Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
(ByVal formPtr As LongPtr, ByVal handler As LongPtr, ByVal c As LongLong, ByVal d As LongLong) As LongLong

' Find an item in a collection by its ObjPtr
Public Function GetItemByPtr(ByRef collection As Variant, ByVal itemPtr As LongPtr) As Variant
    Dim item As Variant
    For Each item In collection
        If ObjPtr(item) = itemPtr Then
            Set GetItemByPtr = item
            Exit Function
        End If
    Next
    Set GetItemByPtr = Nothing
End Function

Private Sub MacDoResize(ByVal formPtr As LongPtr, ByVal Left As Double, ByVal Top As Double, ByVal Width As Double, ByVal Height As Double)
    Dim form As Variant
    Set form = GetItemByPtr(UserForms, formPtr)
    If Not (form Is Nothing) Then
        form.Move Left, Top, Width, Height
        form.Repaint
    End If
End Sub

' caller must ensure that `currentForm` is the current active form
' TODO: Get the current active form programmatically, instead of
Public Sub MakeFormResizable(currentForm As Variant)
    MacMakeFormResizable ObjPtr(currentForm), AddressOf MacDoResize, 0, 0
End Sub
#Else
' Code to make UserForm resizable

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

Public Sub MakeFormResizable(currentForm As UserForm)

  Dim lStyle As Long
  Dim hWnd As Long
  Dim RetVal As Long
  
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

' Functions to get the DPI (not currently used as this is not reliable).
' DLLs and global variables declared/defined at the top

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

 hDC = GetDC(0)
 lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
 ReleaseDC 0, hDC

End Function
#End If

' Path functions
Public Function GetTempPath() As String
    Dim res As String
    res = GetITSetting("Temp Dir", DEFAULT_TEMP_DIR)
    If res = vbNullString Then
        #If Mac Then
            res = MacTempPath()
        #Else
            ' we shouldn't reach this condition on Windows, but just in case
            res = GetITSetting("Abs Temp Dir", DEFAULT_TEMP_DIR)
        #End If
    End If
    res = AddTrailingSlash(res)
    GetTempPath = res
End Function

Public Function GetEditorPath() As String
    Dim res As String
    res = GetITSetting("Editor", DEFAULT_EDITOR)
    GetEditorPath = res
End Function

Public Function IsPathWritable(ByVal TempPath As String) As Boolean
    Dim FilePrefix As String
    FilePrefix = DefaultFilePrefix
    
    Dim Fname As String
    Dim FHdl As Integer
    Fname = TempPath & FilePrefix & ".tmp"
    On Error GoTo TempFolderNotWritable
    FHdl = FreeFile()
    Open Fname For Output Access Write As FHdl
    Print #FHdl, "TESTWRITE"
    Close FHdl
    IsPathWritable = True
    Kill Fname
    
    On Error GoTo 0
    
    Exit Function

TempFolderNotWritable:
    IsPathWritable = False
    MsgBox "The temporary folder " & TempPath & " appears not to be writable." & vbCrLf & _
            "If you're trying to use a relative path, you need to have saved your presentation once."
End Function

Public Function CleanPath(TempPath As String) As String
    If Right$(TempPath, 1) <> PathSep Then
        TempPath = TempPath & PathSep
    End If
    If Left$(TempPath, 1) = "." Then
        Dim sPath As String
        sPath = ActivePresentation.path
        If Len(sPath) > 0 Then
            If Right$(sPath, 1) <> PathSep Then
                sPath = sPath & PathSep
            End If
            TempPath = sPath & TempPath
        Else
            TempPath = PathSep ' This will raise an error in IsPathWritable
        End If
    End If
    CleanPath = TempPath
End Function

Public Function GetExtension(ByVal FileName As String) As String
    GetExtension = LCase$(Right$(FileName, Len(FileName) - InStrRev(FileName, ".")))
End Function

Public Function GetFolderFromPath(strFullPath As String) As String
    GetFolderFromPath = Left(strFullPath, InStrRev(strFullPath, PathSep))
End Function

Public Function BrowseFilePath(Optional ByVal InitialName As String = vbNullString, _
                        Optional ByVal NameFilterDesc As String = vbNullString, _
                        Optional ByVal NameFilterExt As String = vbNullString, _
                        Optional ByVal ButtonNAmeStr As String = vbNullString) As String
    #If Mac Then
        BrowseFilePath = AppleScriptTask("IguanaTex.scpt", "MacChooseFile", InitialName)
        
    #Else
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = InitialName
    fd.Filters.Clear
    If NameFilterExt <> vbNullString Then fd.Filters.Add NameFilterDesc, NameFilterExt, 1
    If ButtonNAmeStr <> vbNullString Then fd.ButtonNAme = ButtonNAmeStr
    
    BrowseFilePath = vbNullString
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            BrowseFilePath = vrtSelectedItem
        Next vrtSelectedItem
    End If

    If BrowseFilePath = vbNullString Then
        BrowseFilePath = InitialName
    End If

    Set fd = Nothing
    #End If
End Function

Public Function BrowseFolderPath(Optional ByVal InitialName As String = vbNullString) As String
    #If Mac Then
        BrowseFolderPath = AppleScriptTask("IguanaTex.scpt", "MacChooseFolder", InitialName)
        
    #Else
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = InitialName
    
    BrowseFolderPath = vbNullString
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            BrowseFolderPath = vrtSelectedItem
        Next vrtSelectedItem
    End If

    If BrowseFolderPath = vbNullString Then
        BrowseFolderPath = InitialName
    End If
    
    Set fd = Nothing
    #End If
End Function

Public Sub OpenURL(ByVal Link As String)
    Dim lSuccess As Long
    #If Mac Then
        lSuccess = CLng(AppleScriptTask("IguanaTex.scpt", "MacExecute", "open " & ShellEscape(Link)))
    #Else
        lSuccess = ShellExecute(0, "Open", Link)
    #End If
    If (lSuccess = 0) Then
        MsgBox "Cannot open " & Link
    End If
End Sub

Public Function ReadAll(FileName As String) As String
    If FileExists(FileName) Then
        #If Mac Then
            ReadAll = Utf8ToString(ReadAllBytes(FileName))
        #Else
            Dim objStream As Object
            Set objStream = CreateObject("ADODB.Stream")
            objStream.Charset = "utf-8"
            objStream.Open
            objStream.LoadFromFile (FileName)
            ReadAll = objStream.ReadText()
        #End If
    Else
        ReadAll = vbNullString
    End If
End Function

Public Function ReadTextFile(ByVal FileName As String) As String
    ' non UTF8 version, no longer used, only for Windows
    Const ForReading As Long = 1
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If FileExists(FileName) Then
        Dim TextFile As Object
        Set TextFile = fs.OpenTextFile(FileName, ForReading)
        ReadTextFile = TextFile.ReadAll
        TextFile.Close
    Else
        ReadTextFile = vbNullString
    End If
    Set fs = Nothing
End Function


Public Sub WriteToFile(ByVal TempPath As String, ByVal FilePrefix As String, ByVal InputText As String)
    #If Mac Then
        Dim fs As New MacFileSystemObject
        If fs.FileExists(TempPath & FilePrefix & ".png") Then
            fs.FindDelete TempPath, FilePrefix + "*.*" 'Make sure we don't keep old files
        End If
    
        ' always use utf-8
    
        Dim fnum As Integer
    
        ' clear file content
        fnum = FreeFile()
        Open TempPath + FilePrefix + ".tex" For Output Access Write As fnum
        Close #fnum
    
        ' write data
        Dim data() As Byte
        data = StringToUtf8(InputText)
    
        fnum = FreeFile()
        Open TempPath + FilePrefix + ".tex" For Binary Access Write As fnum
        Put #fnum, , data
        Close #fnum
    
        Set fs = Nothing
    #Else
        Dim fs As New FileSystemObject
        If fs.FileExists(TempPath & FilePrefix & ".png") Then
            fs.DeleteFile TempPath + FilePrefix + "*.*" 'Make sure we don't keep old files
        End If
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
            .WriteText InputText
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
        
        Set fs = Nothing
    #End If
    
End Sub

Public Function FileExists(ByVal pathname As String) As Boolean
    #If Mac Then
        Dim fs As New MacFileSystemObject
    #Else
        Dim fs As New FileSystemObject
    #End If
    FileExists = fs.FileExists(pathname)
    Set fs = Nothing
End Function


' Functions to clean strings
Public Function RemoveQuotes(Str As String) As String
    If Left$(Str, 1) = """" Then Str = Mid$(Str, 2, Len(Str) - 1)
    If Right$(Str, 1) = """" Then Str = Left$(Str, Len(Str) - 1)
    RemoveQuotes = Str
End Function

Public Function AddTrailingSlash(Str As String) As String
    If Str <> vbNullString And Right$(Str, 1) <> PathSep Then Str = Str & PathSep
    AddTrailingSlash = Str
End Function

' Sanitize booleans to get consistent representation
Public Function BoolToInt(ByVal val As Boolean) As Long
    If val Then
        BoolToInt = 1&
    Else
        BoolToInt = 0&
    End If
End Function

' Wrapper for settings retrieval/storing
Public Function GetITSetting(ByVal Valuename As String, ByVal defaultValue As Variant) As Variant
    GetITSetting = GetRegistryValue(HKEY_CURRENT_USER, RegPath, Valuename, defaultValue)
End Function

Public Sub SetITSetting(ByRef Valuename As String, Valuetype As Long, value As Variant)
    SetRegistryValue HKEY_CURRENT_USER, RegPath, Valuename, Valuetype, value
End Sub


' Define constant arrays as functions
Public Function GetLaTexEngineList() As Variant
    GetLaTexEngineList = Array("pdflatex", "pdflatex", "xelatex", "lualatex", "platex")
End Function

Public Function GetLaTexEngineDisplayList() As Variant
    GetLaTexEngineDisplayList = Array("latex (DVI)", "pdflatex", "xelatex", "lualatex", "platex")
End Function

Public Function GetLatexDVIOptionsList() As Variant
    GetLatexDVIOptionsList = Array("-output-format dvi", "-output-format dvi", "-no-pdf", "-output-format dvi", vbNullString)
End Function

Public Function GetLatexmkPDFOptionsList() As Variant
    GetLatexmkPDFOptionsList = Array("-pdf", "-pdf", "-xelatex", _
        "-lualatex", "-pdfdvi -latex=platex -e ""$dvipdf='dvipdfmx %O %S';$bibtex='pbibtex';""")
End Function

Public Function GetLatexmkDVIOptionsList() As Variant
    GetLatexmkDVIOptionsList = Array("-dvi", "-dvi", "-pdfxe -pdfxelatex=""xelatex --shell-escape %O %S""", _
        "-dvi -pdf- -latex=""dvilualatex --shell-escape %O %S""", "-dvi -latex=platex -e ""$bibtex='pbibtex';""")
End Function

Public Function GetUseDVIList() As Variant
    GetUseDVIList = Array(True, False, False, False, True)
End Function

Public Function GetUsePDFList() As Variant
    GetUsePDFList = Array(False, True, True, True, True)
End Function

Public Function GetBitmapVectorList() As Variant
    GetBitmapVectorList = Array("Bitmap", "Vector")
End Function

Public Function GetVectorOutputTypeDisplayList() As Variant
    #If Mac Then
        GetVectorOutputTypeDisplayList = Array("SVG via DVI w/ dvisvgm", "SVG via PDF w/ dvisvgm")
    #Else
        GetVectorOutputTypeDisplayList = Array("SVG via DVI w/ dvisvgm", "SVG via PDF w/ dvisvgm", "EMF w/ TeX2img", "EMF w/ pdfiumdraw")
    #End If
End Function

Public Function GetVectorOutputTypeList() As Variant
    #If Mac Then
        GetVectorOutputTypeList = Array("dvisvgm", "dvisvgmpdf")
    #Else
        GetVectorOutputTypeList = Array("dvisvgm", "dvisvgmpdf", "tex2img", "pdfiumdraw")
    #End If
End Function

' Shape functions

' Helper for full ungrouping of shapes (handles sub-groups)
Public Sub FullyUngroupShape(NewShape As Shape)
    Dim Shr As ShapeRange
    Dim i As Long
    Dim s As Shape
    If NewShape.Type = msoGroup Then
        Set Shr = NewShape.Ungroup
        For i = 1 To Shr.count
            Set s = Shr.item(i)
            If s.Type = msoGroup Then
                FullyUngroupShape s
            End If
        Next
    End If
End Sub

' Get list of all shapes in a group so that we can re-group them later on
Public Function GetAllShapesInGroup(NewShape As Shape) As Variant
    Dim arr() As Variant
    Dim j As Long
    Dim s As Shape
    For Each s In NewShape.GroupItems
            j = j + 1
            ReDim Preserve arr(1 To j)
            arr(j) = s.Name
    Next
    GetAllShapesInGroup = arr
End Function

' Add picture as shape taking care of not inserting it in empty placeholder
Public Function AddDisplayShape(ByVal path As String, ByVal posX As Single, ByVal posY As Single) As Shape
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
    Set AddDisplayShape = osld.Shapes.AddPicture(path, msoFalse, msoTrue, posX, posY, -1, -1)
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

' Clean up Shape from EMF picture
Public Function ConvertEMF(inSh As Shape, ByVal ScalingX As Single, ByVal ScalingY As Single, _
                                           Optional ByVal posX As Long = -1, Optional ByVal posY As Long = -1, _
                                           Optional ByVal FileType As String = "emf", _
                                           Optional ByVal ConvertLines As Boolean = True, _
                                           Optional ByVal CleanUp As Boolean = True) As Shape
    With inSh
        .ScaleHeight 1#, msoTrue
        .ScaleWidth 1#, msoTrue
        .LockAspectRatio = msoFalse
        .ScaleHeight ScalingY, msoTrue
        .ScaleWidth ScalingX, msoTrue
        .LockAspectRatio = msoTrue
    End With
    
    Dim NewShape As Shape
    ' Get current slide, it will be used to group ranges
    Dim sld As Slide
    Dim SlideIndex As Long
    SlideIndex = ActiveWindow.View.Slide.SlideIndex
    Set sld = ActivePresentation.Slides(SlideIndex)

    ' Convert EMF image to object
    Dim Shr As ShapeRange
    If CleanUp Then
        Set Shr = inSh.Ungroup
        If FileType = "emf" Then
            Set Shr = Shr.Ungroup
            ' Clean up
            Shr.item(1).Delete
            Shr.item(2).Delete
            If Shr(3).GroupItems.count > 2 Then
                Set NewShape = Shr(3)
            Else ' only a single freeform, so not a group
                Set NewShape = Shr(3).GroupItems(2)
            End If
            Shr(3).GroupItems(1).Delete
        ElseIf FileType = "eps" Then
            If CleanUp Then
                Shr.GroupItems(1).Delete
                Shr.GroupItems(1).Delete
            End If
            Set NewShape = Shr.Ungroup.Group
        End If
    Else
        Set NewShape = inSh
    End If
    
    If NewShape.Type = msoGroup Then
    
        Dim arr_group() As Variant
        arr_group = GetAllShapesInGroup(NewShape)
        FullyUngroupShape NewShape
        Set NewShape = sld.Shapes.Range(arr_group).Group
        
        Dim emf_arr() As Variant ' gather all shapes to be regrouped later on
        Dim delete_arr() As Variant ' gather all shapes to be deleted later on
        Dim j_emf As Long, j_delete As Long
        j_emf = 0
        j_delete = 0
        Dim s As Shape
        For Each s In NewShape.GroupItems
            j_emf = j_emf + 1
            ReDim Preserve emf_arr(1 To j_emf)
            If s.Type = msoLine Then
                If ConvertLines And (s.Height > 0 Or s.Width > 0) Then
                    emf_arr(j_emf) = LineToFreeform(s).Name
                    j_delete = j_delete + 1
                    ReDim Preserve delete_arr(1 To j_delete)
                    delete_arr(j_delete) = s.Name
                Else
                    emf_arr(j_emf) = s.Name
                End If
            Else
                emf_arr(j_emf) = s.Name
                If s.Fill.Visible = msoTrue Then
                    s.Line.Visible = msoFalse
                Else
                    s.Line.Visible = msoTrue
                End If
            End If
        Next
        NewShape.Ungroup
        If j_delete > 0 Then
            sld.Shapes.Range(delete_arr).Delete
        End If
        Set NewShape = sld.Shapes.Range(emf_arr).Group
    
    Else
        If NewShape.Type = msoLine Then
            Dim newShapeName As String
            newShapeName = LineToFreeform(NewShape).Name
            NewShape.Delete
            Set NewShape = sld.Shapes(newShapeName)
        Else
            NewShape.Line.Visible = msoFalse
        End If
    End If
    
    NewShape.LockAspectRatio = msoTrue
    If posX <> -1 Then NewShape.Left = posX
    If posY <> -1 Then NewShape.Top = posY
    
    Set ConvertEMF = NewShape
End Function


Public Function convertSVG(inSh As Shape, ByVal ScalingX As Single, ByVal ScalingY As Single, _
                           Optional ByVal posX As Long = -1, Optional ByVal posY As Long = -1) As Shape
    With inSh
        .ScaleHeight 1#, msoTrue
        .ScaleWidth 1#, msoTrue
        .LockAspectRatio = msoFalse
        .ScaleHeight ScalingY, msoTrue
        .ScaleWidth ScalingX, msoTrue
        .LockAspectRatio = msoTrue
    End With
    ' Because we're applying a ribbon function, we need to use selection to keep track of the shape
    Dim Sel As Selection
    inSh.Select
    Call CommandBars.ExecuteMso("SVGEdit")
    Set Sel = Application.ActiveWindow.Selection
    Dim NewShape As Shape
    
    ' Get current slide, it will be used to group ranges
    Dim sld As Slide
    Dim SlideIndex As Long
    SlideIndex = ActiveWindow.View.Slide.SlideIndex
    Set sld = ActivePresentation.Slides(SlideIndex)
    Set NewShape = Sel.ShapeRange(1)
    
    ' For SVG, it looks like there isn't really anything to clean up,
    ' as the group structure is already flat and there are no visible Lines,
    ' but it probably doesn't hurt to have this code either
    If NewShape.Type = msoGroup Then
        Dim arr_group() As Variant
        arr_group = GetAllShapesInGroup(NewShape)
        FullyUngroupShape NewShape
        Set NewShape = sld.Shapes.Range(arr_group).Group
        
        Dim s As Shape
        For Each s In NewShape.GroupItems
            If s.Fill.Visible = msoTrue Then
                 s.Line.Visible = msoFalse
             Else
                s.Line.Visible = msoTrue
             End If
        Next
    Else
        NewShape.Line.Visible = msoFalse
    End If
    NewShape.Select
    
    NewShape.LockAspectRatio = msoTrue
    If posX <> -1 Then NewShape.Left = posX
    If posY <> -1 Then NewShape.Top = posY
    
    Set convertSVG = NewShape
End Function



Private Function LineToFreeform(ByVal s As Shape) As Shape
    Dim t As Double
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
    Dim xs As Double: xs = aC(nBegin, 1)
    Dim ys As Double: ys = aC(nBegin, 2)
    Dim xe As Double: xe = aC(nEnd, 1)
    Dim ye As Double: ye = aC(nEnd, 2)
    
    ' Get unit vector in orthogonal direction
    Dim xd As Double: xd = xe - xs
    Dim yd As Double: yd = ye - ys
    
    Dim s_length As Double: s_length = Sqr(xd * xd + yd * yd)
    Dim n_x As Double
    Dim n_y As Double
    If s_length > 0 Then
        n_x = -yd / s_length
        n_y = xd / s_length
    Else
        n_x = 0
        n_y = 0
    End If
    
    Dim x1 As Double: x1 = xs + n_x * t / 2
    Dim y1 As Double: y1 = ys + n_y * t / 2
    Dim x2 As Double: x2 = xe + n_x * t / 2
    Dim y2 As Double: y2 = ye + n_y * t / 2
    Dim x3 As Double: x3 = xe - n_x * t / 2
    Dim y3 As Double: y3 = ye - n_y * t / 2
    Dim x4 As Double: x4 = xs - n_x * t / 2
    Dim y4 As Double: y4 = ys - n_y * t / 2
        
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


Public Function ArrayLength(arr As Variant) As Long
    On Error GoTo handler
    ArrayLength = UBound(arr) + 1
    Exit Function
handler:
    ArrayLength = 0
End Function


Function RoundUp(ByVal value As Double) As Variant
    If Int(value) = value Then
        RoundUp = value
    Else
        RoundUp = Int(value) + 1
    End If
End Function


Public Function ShellEscape(Str As String) As String
    #If Mac Then
        ShellEscape = "'" & Replace(Replace(Str, "\", "\\"), "'", "'\''") & "'"
    #Else
        ShellEscape = """" & Str & """"
    #End If
End Function

Sub ShowError(ErrorMessage As String, LastCommand As String, _
              Optional FormTitle As String = "Error while running process", _
              Optional CommandPrompt As String = "Last command:", _
              Optional CloseButtonCaption As String = "Close")
    Dim myForm As ErrorForm
    Set myForm = New ErrorForm
    myForm.Caption = FormTitle
    myForm.LabelLastCommandPrompt.Caption = CommandPrompt
    myForm.LabelError.Caption = ErrorMessage
    myForm.LabelCommand.Caption = LastCommand
    myForm.CloseErrorButton.Caption = CloseButtonCaption
    myForm.Show
    Unload myForm
    Set myForm = Nothing
End Sub
