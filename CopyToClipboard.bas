Attribute VB_Name = "CopyToClipboard"
Option Explicit

#If VBA7 Then
Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
  ByVal dwBytes As LongPtr) As LongPtr
Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
  ByVal lpString2 As Any) As Long
Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat _
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

Sub ClipBoard_SetData(MyString As String)
'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

#If VBA7 Then
   Dim hGlobalMemory As LongPtr
   Dim lpGlobalMemory As LongPtr
   Dim hClipMemory As LongPtr
#Else
   Dim hGlobalMemory As Long
   Dim lpGlobalMemory As Long
   Dim hClipMemory As Long
#End If

Dim x As Long

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
    Exit Sub
  End If

'Clear the Clipboard.
  x = EmptyClipboard()

'Copy the data to the Clipboard.
  hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
  If CloseClipboard() = 0 Then
    MsgBox "Could not close Clipboard."
  End If

End Sub


Function Clipboard(Optional StoreText As String) As String

#If Mac Then
    ' In old forums (e.g., https://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-mso_mac-mso_mac2011/read-clipboard-contents-in-vba-excel-for-mac-not/4485ab02-dced-4bf2-853d-fc667f5782a8),
    ' people complain this doesn't work properly, but it seems to be working in my version of PowerPoint (Office365).
    'Dim myData As DataObject
    'Set myData = New DataObject
    'myData.SetText StoreText
    'myData.PutInClipboard
    AppleScriptTask "IguanaTex.scpt", "MacSetClipboard", StoreText
#Else
    'PURPOSE: Read/Write to Clipboard
    'Source: ExcelHero.com (Daniel Ferry)
    ' https://stackoverflow.com/questions/14219455/excel-vba-code-to-copy-a-specific-string-to-clipboard/60896244#60896244
    
    Dim x As Variant
    
    'Store as variant for 64-bit VBA support
      x = StoreText
    
    'Create HTMLFile Object
    With CreateObject("htmlfile")
      With .parentWindow.clipboardData
        Select Case True
          Case Len(StoreText)
            'Write to the clipboard
              .setData "text", x
          Case Else
            'Read from the clipboard (no variable passed through)
              Clipboard = .GetData("text")
        End Select
      End With
    End With
#End If
End Function

