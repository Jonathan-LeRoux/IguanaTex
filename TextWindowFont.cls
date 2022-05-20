VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextWindowFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Replacement for TextBox.Font

Option Explicit

#If Mac Then
Private Declare PtrSafe Function TWGetFontSize _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As Double

Private Declare PtrSafe Function TWSetFontSize _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal value As Double, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Const UNUSED As LongLong = 0

' members
Public mHandle As LongLong

' properties
Public Property Get Size() As Double
    Size = CCur(TWGetFontSize(mHandle, UNUSED, UNUSED, UNUSED))
End Property

Public Property Let Size(value As Double)
    TWSetFontSize mHandle, val(value), UNUSED, UNUSED
End Property

' methods
Public Sub Init(Handle As LongLong)
    mHandle = Handle
End Sub

#End If


