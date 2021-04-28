Attribute VB_Name = "MacUtils"
Option Explicit
Option Private Module

Public Const gUserFormResizeFactor As Double = 1.333333

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
    Dim ctrl As Control
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








