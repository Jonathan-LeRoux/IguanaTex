Attribute VB_Name = "ShellWait"
Option Explicit

#If Mac Then
Public Function Execute(ByVal CommandLine As String, StartupDir As String, Optional debugMode As Boolean = False, Optional WaitTime As Long = -1) As Long
    Dim TeXExePath As String
    TeXExePath = GetITSetting("TeXExePath", DEFAULT_TEX_EXE_PATH)
    Dim TeXExtraPath As String
    TeXExtraPath = GetITSetting("TeXExtraPath", DEFAULT_TEX_EXTRA_PATH)
    If TeXExtraPath <> vbNullString Then
        TeXExtraPath = ":" & TeXExtraPath
    End If
    If debugMode Then
        ShowError vbNullString, CommandLine, "Debug mode", "Next command:", "Continue"
    End If
    Execute = CLng(AppleScriptTask("IguanaTex.scpt", "MacExecute", _
        "export PATH=" & ShellEscape(TeXExePath) & ShellEscape(TeXExtraPath) & """:$PATH""" & " && " & _
        "cd " & ShellEscape(StartupDir) & " && " & _
        CommandLine))
End Function

#Else
' Portions of code below taken from:
' http://www.mvps.org/access/api/api0004.htm
' Courtesy of Terry Kreft

Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadId As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
Private Declare PtrSafe Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long
    
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long
    
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
    
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Public Declare PtrSafe Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal Operation As String, _
  ByVal FileName As String, Optional ByVal Parameters As String, _
  Optional ByVal directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

#Else
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long
    
Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
    
Private Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    
Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal Operation As String, _
  ByVal Filename As String, Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
#End If

    
Public Function ShellWait(pathname As String, Optional StartupDir As String, Optional WindowStyle As Long, Optional WaitTime As Long = -1) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long
    Dim exitcode As Long
    Dim lastError As Long
    Dim retWait As Long
    
    ' Initialize the STARTUPINFO structure:
    With start
        .cb = Len(start)
        If Not IsMissing(WindowStyle) Then
            .dwFlags = STARTF_USESHOWWINDOW
            .wShowWindow = WindowStyle
        End If
    End With
    Dim sdir As String
    If IsMissing(StartupDir) Then
        sdir = vbNullString
    Else
        sdir = StartupDir
    End If

    ' Start the shelled application:
    ret& = CreateProcessA(0&, pathname, 0&, 0&, 1&, _
            NORMAL_PRIORITY_CLASS, 0&, sdir, start, proc)
    lastError& = GetLastError()
    If (ret& = 0) Then
        MsgBox "Could not start process: '" & pathname & "'. GetLastError returned " & Str$(lastError&)
        ShellWait = 1
        Exit Function
    End If
        
    ' Wait for the shelled application to finish:
    If WaitTime > 0 Then
        retWait& = WaitForSingleObject(proc.hProcess, WaitTime)
    Else
        retWait& = WaitForSingleObject(proc.hProcess, INFINITE)
    End If
    ' Get return value
    exitcode& = 1234
    ret& = GetExitCodeProcess(proc.hProcess, exitcode&)
    If (ret& = 0) Then
        lastError& = GetLastError()
        MsgBox "GetExitCodeProcess returned " + Str$(ret&) + ", GetLastError returned " + Str$(lastError&)
    End If
    ' Tidy up if time out
    If (retWait& = 258) Then
        ret& = TerminateProcess(proc.hProcess, 0)
    End If
    ' Close handle
    ret& = CloseHandle(proc.hProcess)
    ShellWait = exitcode&
End Function

Public Function Execute(CommandLine As String, StartupDir As String, Optional debugMode As Boolean = False, Optional WaitTime As Long = -1) As Long
    Dim RetVal As Long
    If debugMode Then
        ' Clipboard CommandLine
        ' MsgBox CommandLine, , StartupDir
        ShowError vbNullString, CommandLine, "Debug mode", "Next command:", "Continue"
        RetVal = ShellWait(CommandLine, StartupDir, 1&, WaitTime)
    Else
        RetVal = ShellWait(CommandLine, StartupDir, , WaitTime)
    End If
    Execute = RetVal
End Function
#End If ' Mac


