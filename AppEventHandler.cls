VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@MemberAttribute VB_VarHelpID, -1
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WindowBeforeDoubleClick(ByVal Sel As Selection, ByRef Cancel As Boolean)
    If (TryEditLatexEquation()) Then
        Cancel = True
    End If
End Sub
