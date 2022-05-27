Attribute VB_Name = "MacUtils"
Option Explicit
Option Private Module

Public Const gUserFormResizeFactor As Double = 1.333333

#If Mac Then
' macOS virtual keycodes. See <HIToolbox/Events.h>
' or https://developer.mozilla.org/en-US/docs/Web/API/UI_Events/Keyboard_event_code_values#code_values_on_mac
Private Const kVK_ANSI_A As LongLong = &H0
Private Const kVK_ANSI_Z As LongLong = &H6
Private Const kVK_ANSI_X As LongLong = &H7
Private Const kVK_ANSI_C As LongLong = &H8
Private Const kVK_ANSI_V As LongLong = &H9

Private Const NSEventModifierFlagShift As LongLong = &H20000
Private Const NSEventModifierFlagCommand As LongLong = &H100000

Private Declare PtrSafe Function MacEnableCopyPaste_Native _
 Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
 Alias "MacEnableCopyPaste" _
(ByVal formPtr As LongPtr, ByVal handler As LongPtr, ByVal c As LongLong, ByVal d As LongLong) As LongLong
Private Declare PtrSafe Function MacEnableAccelerators_Native _
 Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
 Alias "MacEnableAccelerators" _
(ByVal formPtr As LongPtr, ByVal handler As LongPtr, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Sub MacDoCopyPaste(ByVal formPtr As LongPtr, ByVal keyCode As LongLong, ByVal modifierFlags As LongLong)
    Dim form As Variant
    Set form = GetItemByPtr(UserForms, formPtr)
    If form Is Nothing Then
        Exit Sub
    End If

    If keyCode = kVK_ANSI_Z And modifierFlags = NSEventModifierFlagCommand Then
        form.UndoAction
        Exit Sub
    ElseIf keyCode = kVK_ANSI_Z And modifierFlags = NSEventModifierFlagCommand + NSEventModifierFlagShift Then
        form.RedoAction
        Exit Sub
    End If

    Dim ctrl As control
    Set ctrl = form.ActiveControl
    Do
        If TypeOf ctrl Is Frame Then
            Set ctrl = ctrl.ActiveControl
        ElseIf TypeOf ctrl Is MultiPage Then
            Set ctrl = ctrl.Pages(ctrl.value).ActiveControl
        Else
            Exit Do
        End If
    Loop

    If Not (TypeOf ctrl Is TextBox) Then
        Exit Sub
    End If

    If keyCode = kVK_ANSI_A And modifierFlags = NSEventModifierFlagCommand Then
        ctrl.SelStart = 0
        ctrl.SelLength = Len(ctrl)
    ElseIf keyCode = kVK_ANSI_C And modifierFlags = NSEventModifierFlagCommand Then
        ctrl.Copy
    ElseIf keyCode = kVK_ANSI_V And modifierFlags = NSEventModifierFlagCommand Then
        ctrl.Paste
    ElseIf keyCode = kVK_ANSI_X And modifierFlags = NSEventModifierFlagCommand Then
        ctrl.Cut
    End If
End Sub

Private Sub MacDoAccelerator(ByVal formPtr As LongPtr, ByVal asciiCode As LongLong)
    Dim form As Variant
    Set form = GetItemByPtr(UserForms, formPtr)
    If form Is Nothing Then
        Debug.Print "Form Not Found"
        Exit Sub
    End If

    Dim ch As String
    ch = UCase(Chr(CLng(asciiCode)))
    If ch = vbNullString Then
        Exit Sub
    End If

    Dim ctrl As control
    On Error Resume Next
    For Each ctrl In form.Controls
        If TypeOf ctrl Is MultiPage Then
            Dim page As page
            For Each page In ctrl.Pages
                If UCase(Left(page.Accelerator, 1)) = ch Then
                    ctrl.value = page.index
                    Exit Sub
                End If
            Next
        ElseIf TypeOf ctrl Is CheckBox Or TypeOf ctrl Is CommandButton Then
            If UCase(Left(ctrl.Accelerator, 1)) = ch Then
                If TypeOf ctrl Is CheckBox Then
                    ctrl.value = IsNull(ctrl.value) Or Not ctrl.value
                End If
                CallByName form, ctrl.Name & "_Click", VbMethod
                Exit Sub
            End If
        End If
    Next
End Sub

Public Sub MacEnableCopyPaste(ByRef currentForm As Variant)
    MacEnableCopyPaste_Native ObjPtr(currentForm), AddressOf MacDoCopyPaste, 0, 0
End Sub

Public Sub MacEnableAccelerators(ByRef currentForm As Variant)
    MacEnableAccelerators_Native ObjPtr(currentForm), AddressOf MacDoAccelerator, 0, 0
End Sub
#End If

Public Function ReadAllBytes(FileName As String) As Byte()
    Dim fnum As Integer
    fnum = FreeFile()
    Open FileName For Binary Access Read As fnum

    Dim length As Long
    length = FileLen(FileName)

    Dim data() As Byte
    If length <> 0 Then
        ReDim data(length - 1)
        Get #fnum, , data
    End If

    Close #fnum

    ReadAllBytes = data
End Function

#If Mac Then
Public Function MacTempPath() As String
    MacTempPath = MacScript("POSIX path of (path to temporary items)")
End Function

Public Function MacChooseFileOfType(typeStr As String) As String
    MacChooseFileOfType = AppleScriptTask("IguanaTex.scpt", "MacChooseFileOfType", typeStr)
End Function

Public Function MacChooseApp(defaultValue As String) As String
    MacChooseApp = AppleScriptTask("IguanaTex.scpt", "MacChooseApp", defaultValue)
End Function
#End If

Sub ResizeUserForm(frm As Object, Optional dResizeFactor As Double = 0#)
'Created by Jon Peltier
    Dim ctrl As control
    Dim sColWidths As String
    Dim vColWidths As Variant
    Dim iCol As Long
    
    If dResizeFactor = 0 Then dResizeFactor = gUserFormResizeFactor
    With frm
        '.Resize = True
        .Height = .Height * dResizeFactor
        .Width = .Width * dResizeFactor
        
        For Each ctrl In frm.Controls
            With ctrl
                .Height = .Height * dResizeFactor
                .Width = .Width * dResizeFactor
                .Left = .Left * dResizeFactor
                .Top = .Top * dResizeFactor
                On Error Resume Next
                .Font.Size = .Font.Size * dResizeFactor
                On Error GoTo 0
                
                ' multi column listboxes, comboboxes
                Select Case TypeName(ctrl)
                    Case "ListBox", "ComboBox"
                        If ctrl.ColumnCount > 1 Then
                        sColWidths = ctrl.ColumnWidths
                        vColWidths = Split(sColWidths, ";")
                        For iCol = LBound(vColWidths) To UBound(vColWidths)
                        vColWidths(iCol) = val(vColWidths(iCol)) * dResizeFactor
                        Next
                        sColWidths = Join(vColWidths, ";")
                        ctrl.ColumnWidths = sColWidths
                    End If
                End Select
            End With
        Next
    End With
End Sub








