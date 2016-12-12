Attribute VB_Name = "MouseWheel"
' From https://social.msdn.microsoft.com/Forums/en-US/7d584120-a929-4e7c-9ec2-9998ac639bea/mouse-scroll-in-userform-listbox-in-excel-2010?forum=isvvba
'
'Enables mouse wheel scrolling in controls
Option Explicit

#If Win64 Then
    Private Type POINTAPI
       XY As LongLong
    End Type
#Else
    Private Type POINTAPI
           X As Long
           Y As Long
    End Type
#End If

Private Type MOUSEHOOKSTRUCT
    Pt As POINTAPI
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" _
                                            Alias "FindWindowA" ( _
                                                            ByVal lpClassName As String, _
                                                            ByVal lpWindowName As String) As Long ' not sure if this should be LongPtr
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" _
                                            Alias "GetWindowLongPtrA" ( _
                                                            ByVal hWnd As LongPtr, _
                                                            ByVal nIndex As Long) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLong Lib "user32" _
                                            Alias "GetWindowLongA" ( _
                                                            ByVal hWnd As LongPtr, _
                                                            ByVal nIndex As Long) As LongPtr
    #End If
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" _
                                            Alias "SetWindowsHookExA" ( _
                                                            ByVal idHook As Long, _
                                                            ByVal lpfn As LongPtr, _
                                                            ByVal hmod As LongPtr, _
                                                            ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
                                                            ByVal hHook As LongPtr, _
                                                            ByVal nCode As Long, _
                                                            ByVal wParam As LongPtr, _
                                                           lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
                                                            ByVal hHook As LongPtr) As LongPtr ' MAYBE Long
    #If Win64 Then
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
                                                            ByVal Point As LongLong) As LongPtr    '
    #Else
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
                                                            ByVal xPoint As Long, _
                                                            ByVal yPoint As Long) As LongPtr    '
    #End If
    Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
                                                            ByRef lpPoint As POINTAPI) As LongPtr   'MAYBE Long
#Else
    Private Declare Function FindWindow Lib "user32" _
                                            Alias "FindWindowA" ( _
                                                            ByVal lpClassName As String, _
                                                            ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32.dll" _
                                            Alias "GetWindowLongA" ( _
                                                            ByVal hWnd As Long, _
                                                            ByVal nIndex As Long) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" _
                                            Alias "SetWindowsHookExA" ( _
                                                            ByVal idHook As Long, _
                                                            ByVal lpfn As Long, _
                                                            ByVal hmod As Long, _
                                                            ByVal dwThreadId As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" ( _
                                                            ByVal hHook As Long, _
                                                            ByVal nCode As Long, _
                                                            ByVal wParam As Long, _
                                                           lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
                                                            ByVal hHook As Long) As Long
    Private Declare Function WindowFromPoint Lib "user32" ( _
                                                            ByVal xPoint As Long, _
                                                            ByVal yPoint As Long) As Long
    Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                                            ByRef lpPoint As POINTAPI) As Long
#End If

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)
'Private Const WM_KEYDOWN As Long = &H100
'Private Const WM_KEYUP As Long = &H101
'Private Const VK_UP As Long = &H26
'Private Const VK_DOWN As Long = &H28
'Private Const WM_LBUTTONDOWN As Long = &H201
Dim n As Long
Private mCtl As MSForms.control
Private mbHook As Boolean
#If VBA7 Then
    Private mLngMouseHook As LongPtr
    Private mListBoxHwnd As LongPtr
#Else
    Private mLngMouseHook As Long
    Private mListBoxHwnd As Long
#End If
     
Sub HookListBoxScroll(frm As Object, ctl As MSForms.control)
    Dim tPT As POINTAPI
    #If VBA7 Then
        Dim lngAppInst As LongPtr
        Dim hwndUnderCursor As LongPtr
    #Else
        Dim lngAppInst As Long
        Dim hwndUnderCursor As Long
    #End If
    GetCursorPos tPT
    #If Win64 Then
        hwndUnderCursor = WindowFromPoint(tPT.XY)
    #Else
        hwndUnderCursor = WindowFromPoint(tPT.X, tPT.Y)
    #End If
    If Not frm.ActiveControl Is ctl Then
           ctl.SetFocus
    End If
    If mListBoxHwnd <> hwndUnderCursor Then
        UnhookListBoxScroll
        Set mCtl = ctl
        mListBoxHwnd = hwndUnderCursor
        #If Win64 Then
            lngAppInst = GetWindowLongPtr(mListBoxHwnd, GWL_HINSTANCE)
        #Else
            lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
        #End If
        ' PostMessage mListBoxHwnd, WM_LBUTTONDOWN, 0&, 0&
        If Not mbHook Then
            mLngMouseHook = SetWindowsHookEx( _
                                            WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
            mbHook = mLngMouseHook <> 0
        End If
    End If
End Sub

Sub UnhookListBoxScroll()
    If mbHook Then
        Set mCtl = Nothing
        UnhookWindowsHookEx mLngMouseHook
        mLngMouseHook = 0
        mListBoxHwnd = 0
        mbHook = False
    End If
End Sub
#If VBA7 Then
Private Function MouseProc( _
                        ByVal nCode As Long, ByVal wParam As Long, _
                        ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
#Else
Private Function MouseProc( _
                        ByVal nCode As Long, ByVal wParam As Long, _
                        ByRef lParam As MOUSEHOOKSTRUCT) As Long
#End If
    Dim idx As Long
    On Error GoTo errH
    If (nCode = HC_ACTION) Then
        #If Win64 Then
            If WindowFromPoint(lParam.Pt.XY) = mListBoxHwnd Then
                If wParam = WM_MOUSEWHEEL Then
                    MouseProc = True
'                        If lParam.hWnd > 0 Then
'                            postMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
'                        Else
'                            postMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
'                        End If
'                        postMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
                    If TypeOf mCtl Is Frame Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    ElseIf TypeOf mCtl Is UserForm Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    ElseIf TypeOf mCtl Is TextBox Then
                        If lParam.hWnd > 0 Then idx = -3 Else idx = 3
                        idx = idx + mCtl.CurLine
                        If idx < 0 Then idx = 0
                        If idx > mCtl.LineCount + 1 Then idx = mCtl.LineCount + 1
                        mCtl.CurLine = idx
                    Else
                        If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                        idx = idx + mCtl.ListIndex
                        If idx >= 0 Then mCtl.ListIndex = idx
                    End If
                Exit Function
                End If
            Else
                UnhookListBoxScroll
            End If
        #Else
            If WindowFromPoint(lParam.Pt.X, lParam.Pt.Y) = mListBoxHwnd Then
                If wParam = WM_MOUSEWHEEL Then
                    MouseProc = True
'                        If lParam.hWnd > 0 Then
'                            postMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
'                        Else
'                            postMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
'                        End If
'                        postMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
                    If TypeOf mCtl Is Frame Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    ElseIf TypeOf mCtl Is UserForm Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    Else
                        If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                        idx = idx + mCtl.ListIndex
                        If idx >= 0 Then mCtl.ListIndex = idx
                    End If
                    Exit Function
                End If
            Else
                UnhookListBoxScroll
            End If
        #End If
    End If
    MouseProc = CallNextHookEx( _
                            mLngMouseHook, nCode, wParam, ByVal lParam)
    Exit Function
errH:
    UnhookListBoxScroll
End Function
'#Else
'    Private Function MouseProc( _
'                            ByVal nCode As Long, ByVal wParam As Long, _
'                            ByRef lParam As MOUSEHOOKSTRUCT) As Long
'        Dim idx As Long
'        On Error GoTo errH
'        If (nCode = HC_ACTION) Then
'            If WindowFromPoint(lParam.Pt.X, lParam.Pt.Y) = mListBoxHwnd Then
'                If wParam = WM_MOUSEWHEEL Then
'                    MouseProc = True
''                    If lParam.hWnd > 0 Then
''                    postMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
''                    Else
''                    postMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
''                    End If
''                    postMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
'                    If TypeOf mCtl Is Frame Then
'                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
'                        idx = idx + mCtl.ScrollTop
'                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
'                            mCtl.ScrollTop = idx
'                        End If
'                    ElseIf TypeOf mCtl Is UserForm Then
'                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
'                        idx = idx + mCtl.ScrollTop
'                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
'                            mCtl.ScrollTop = idx
'                        End If
'                    Else
'                        If lParam.hWnd > 0 Then idx = -1 Else idx = 1
'                        idx = idx + mCtl.ListIndex
'                        If idx >= 0 Then mCtl.ListIndex = idx
'                    End If
'                    Exit Function
'                End If
'            Else
'                UnhookListBoxScroll
'            End If
'        End If
'        MouseProc = CallNextHookEx( _
'        mLngMouseHook, nCode, wParam, ByVal lParam)
'        Exit Function
'errH:
'        UnhookListBoxScroll
'    End Function
'#End If

