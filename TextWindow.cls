VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Replacement for TextBox

Option Explicit

#If Mac Then
Private Declare PtrSafe Function TWInit _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   () As LongLong

Private Declare PtrSafe Function TWTerm _
 Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
(ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWSetResizeTarget _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWResize _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWShow _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWHide _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWFocus _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWGetByteLength _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWGetBytes _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal buffer As LongPtr, ByVal length As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWSetBytes _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal data As LongPtr, ByVal length As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWGetSelStart _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWSetSelStart _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal value As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWGetWordWrap _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWSetWordWrap _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal value As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Const UNUSED As LongLong = 0

' members
Private mHandle As LongLong
Private mFont As New TextWindowFont

' properties
Public Property Get Font() As TextWindowFont
    Set Font = mFont
End Property

Public Property Get Text() As String
    Dim length As LongLong
    length = TWGetByteLength(mHandle, UNUSED, UNUSED, UNUSED)
    
    If length <= 0 Then
        Text = ""
        Exit Property
    End If

    Dim bytes() As Byte
    Dim usedLength As LongLong
    ReDim bytes(CLng(length - 1))
    usedLength = TWGetBytes(mHandle, VarPtr(bytes(0)), length, UNUSED)

    If usedLength <= 0 Then
        Text = ""
        Exit Property
    End If

    ReDim Preserve bytes(CLng(usedLength - 1))
    Text = bytes
End Property

Public Property Let Text(value As String)
    Dim bytes() As Byte
    Dim length As LongLong
    bytes = value
    length = ArrayLength(bytes)
    If length > 0 Then
         TWSetBytes mHandle, VarPtr(bytes(0)), length, UNUSED
    Else
         TWSetBytes mHandle, 0, 0, UNUSED
    End If
End Property

Public Property Get SelStart() As Integer
    SelStart = CInt(TWGetSelStart(mHandle, UNUSED, UNUSED, UNUSED))
End Property

Public Property Let SelStart(value As Integer)
    TWSetSelStart mHandle, CLngLng(value), UNUSED, UNUSED
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = TWGetWordWrap(mHandle, UNUSED, UNUSED, UNUSED)
End Property

Public Property Let WordWrap(value As Boolean)
    TWSetWordWrap mHandle, value, UNUSED, UNUSED
End Property

' constructor/destructor
Private Sub Class_Initialize()
    mHandle = TWInit
    mFont.Init mHandle
End Sub

Private Sub Class_Terminate()
    TWTerm mHandle, UNUSED, UNUSED, UNUSED
End Sub

' methods
Public Sub Show()
    TWShow mHandle, UNUSED, UNUSED, UNUSED
End Sub

Public Sub Hide()
    TWHide mHandle, UNUSED, UNUSED, UNUSED
End Sub

Public Sub SetResizeTarget(target As control, parentForm As UserForm)
    Dim backup As control
    Set backup = parentForm.ActiveControl
    target.SetFocus
    TWSetResizeTarget mHandle, UNUSED, UNUSED, UNUSED
    If Not backup Is Nothing Then
        backup.SetFocus
    End If
End Sub

Public Sub ResizeAsTarget()
    TWResize mHandle, UNUSED, UNUSED, UNUSED
End Sub

Public Sub SetFocus()
    TWFocus mHandle, UNUSED, UNUSED, UNUSED
End Sub

#End If


