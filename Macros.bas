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
 
 Private Declare PtrSafe Function GetDC Lib "user32" _
    (ByVal hwnd As Long) As Long

 Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long

 Private Declare PtrSafe Function ReleaseDC Lib "user32" _
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
   
 Private Declare Function GetDC Lib "user32" _
    (ByVal hwnd As Long) As Long

 Private Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long

 Private Declare Function ReleaseDC Lib "user32" _
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
    
    LatexForm.textboxSize.Visible = True
    LatexForm.Label2.Visible = True
    LatexForm.Label3.Visible = True
    If IsEmpty(LatexForm.textboxSize.Text) Then
        LatexForm.textboxSize.Text = "20"
    End If
       
    LatexForm.ButtonRun.Caption = "Create"
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
    
    If Sel.Type = ppSelectionShapes Then
        ' Attempt to deal with the case of 1 object inside a group
        If Sel.ShapeRange.Type = msoGroup Then
            If Sel.ChildShapeRange.count = 1 Then
                With Sel.ChildShapeRange.Tags
                    For i = 1 To .count
                        If (.name(i) = "LATEXADDIN") Then
                            Load LatexForm
                            LatexForm.textboxSize.Visible = False
                            LatexForm.Label2.Visible = False
                            LatexForm.Label3.Visible = False
    
                            LatexForm.TextBox1.Text = .Value(i)
                            LatexForm.ButtonRun.Caption = "Modify"
                            
                            For j = 1 To .count
                                If (.name(j) = "IGUANATEXSIZE") Then
                                    LatexForm.textboxSize.Text = .Value(j)
                                End If
                            Next
                            LatexForm.textboxSize.Enabled = False
                            LatexForm.Show
                            TryEditLatexEquation = True
                            Exit Function
                        End If
                    Next
                End With
            End If
        ' Now the non-group case: only a single object can be selected
        ElseIf Sel.ShapeRange.count = 1 Then
            With Sel.ShapeRange.Tags
                For i = 1 To .count
                    If (.name(i) = "LATEXADDIN") Then
                        Load LatexForm
                        LatexForm.textboxSize.Visible = False
                        LatexForm.Label2.Visible = False
                        LatexForm.Label3.Visible = False
                        
                        LatexForm.TextBox1.Text = .Value(i)
                        LatexForm.ButtonRun.Caption = "Modify"
                        
                        For j = 1 To .count
                            If (.name(j) = "IGUANATEXSIZE") Then
                                LatexForm.textboxSize.Text = .Value(j)
                            End If
                        Next
                        LatexForm.textboxSize.Enabled = False
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
