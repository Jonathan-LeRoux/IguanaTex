Attribute VB_Name = "ExportVBA"
Option Explicit
' Modified from https://gist.github.com/steve-jansen/7589478 to work in PowerPoint
'
' PowerPoint macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the PowerPoint setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    Const Module As Long = 1
    Const ClassModule As Long = 2
    Const form As Long = 3
    Const Document As Long = 100
    Const Padding As Long = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim Extension As String
    #If Mac Then
        Dim fs As New MacFileSystemObject
    #Else
        Dim fs As New FileSystemObject
    #End If
    
    Dim MyPath As String
    MyPath = ActivePresentation.FullName
    directory = Left$(MyPath, InStrRev(MyPath, ".") - 1) & "_VBA"
    count = 0
    
    If Not fs.FolderExists(directory) Then
        fs.CreateFolder directory
    End If
    Set fs = Nothing
    
    For Each VBComponent In ActivePresentation.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                Extension = ".cls"
            Case form
                Extension = ".frm"
            Case Module
                Extension = ".bas"
            Case Else
                Extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & PathSep & VBComponent.Name & Extension
        VBComponent.Export path
        
        If Err.Number <> 0 Then
            MsgBox "Failed to export " & VBComponent.Name & " to " & path, vbCritical
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    MsgBox "Successfully exported " & CStr(count) & " VBA files to " & directory
    
End Sub
