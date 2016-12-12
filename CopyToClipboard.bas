Attribute VB_Name = "CopyToClipboard"
#If VBA7 Then
Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
  ByVal dwBytes As LongPtr) As LongPtr
Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
  ByVal lpString2 As Any) As Long
Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat _
  As Long, ByVal hMem As LongPtr) As LongPtr
#Else
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
  ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "User32" () As Long
Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "User32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
  ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
  As Long, ByVal hMem As Long) As Long
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Function ClipBoard_SetData(MyString As String)
'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

#If VBA7 Then
   Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr, hClipMemory As LongPtr
#Else
   Dim hGlobalMemory As Long, lpGlobalMemory As Long, hClipMemory As Long
#End If

Dim X As Long

'Allocate moveable global memory
  hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

'Lock the block to get a far pointer to this memory.
  lpGlobalMemory = GlobalLock(hGlobalMemory)

'Copy the string to this global memory.
  lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

'Unlock the memory.
  If GlobalUnlock(hGlobalMemory) <> 0 Then
    MsgBox "Could not unlock memory location. Copy aborted."
    GoTo OutOfHere2
  End If

'Open the Clipboard to copy data to.
  If OpenClipboard(0&) = 0 Then
    MsgBox "Could not open the Clipboard. Copy aborted."
    Exit Function
  End If

'Clear the Clipboard.
  X = EmptyClipboard()

'Copy the data to the Clipboard.
  hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
  If CloseClipboard() = 0 Then
    MsgBox "Could not close Clipboard."
  End If

End Function

