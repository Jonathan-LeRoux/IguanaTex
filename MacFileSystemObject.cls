VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacFileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Partial replacement (in Mac) for Scripting.FileSystemObject

Option Explicit
#If Mac Then

Public Function FileExists(pathname As String) As Boolean
    FileExists = (Dir(pathname) <> "")
End Function

Public Function FolderExists(pathname As String) As Boolean
    FolderExists = (Dir(pathname) <> "")
End Function

Public Sub CreateFolder(pathname As String)
    MkDir pathname
End Sub

'Public Sub CopyFile(pathfrom As String, pathto As String)
'#If Mac Then
'    cp pathfrom pathto
'#Else
'    fs.CopyFile pathfrom, pathto
'#End If
'End Sub

' calls `find <dirname> -name <pattern> -delete`
Public Sub FindDelete(dirname As String, pattern As String)
    AppleScriptTask "IguanaTex.scpt", "MacExecute", "find " & ShellEscape(dirname) & " -name " & ShellEscape(pattern) & " -delete"
End Sub

' does not support path globbing. use `FindDelete` instead.
Public Sub DeleteFile(pathname As String)
    If FileExists(pathname) Then
        ' remove readonly attribute if set
        SetAttr pathname, vbNormal
        Kill pathname
    End If
End Sub
#End If
'
'Private Sub Class_Initialize()
'#If Mac Then
'    '
'#Else
'    Dim fs As New Scripting.FileSystemObject
'#End If
'End Sub
'
'Private Sub Class_Terminate()
'    Set fs = Nothing
'End Sub
