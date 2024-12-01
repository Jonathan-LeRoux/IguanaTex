Attribute VB_Name = "RegistryAccess"
' Portions taken from:
' http://www.kbalertz.com/kb_145679.aspx
   
Option Explicit

Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

#If Mac Then
    Public Function GetRegistryValue(Hive, Keyname, Valuename, DefaultValue)
        If Hive <> HKEY_CURRENT_USER Then
            MsgBox "GetRegistryValue with Hive other than HKEY_CURRENT_USER is not implemented. return defaultValue."
            GetRegistryValue = DefaultValue
            Exit Function
        End If
        
        Dim Str As String
        
        Str = GetSetting("IguanaTex", Keyname, Valuename, "")
        If Str = "" Then
            GetRegistryValue = DefaultValue
        Else
            Dim sp() As String
            sp = Split(Str, ":", 2)
            If UBound(sp) + 1 < 2 Then
                GetRegistryValue = DefaultValue
            ElseIf sp(0) = "sz" Then
                GetRegistryValue = sp(1)
            ElseIf sp(0) = "dword" Then
                GetRegistryValue = CLng(sp(1))
            Else
                GetRegistryValue = DefaultValue
            End If
        End If
    End Function
    
    
    Public Sub SetRegistryValue(Hive, ByRef Keyname As String, ByRef Valuename As String, _
    Valuetype As Long, value As Variant)
        If Hive <> HKEY_CURRENT_USER Then
            MsgBox "SetRegistryValue with Hive other than HKEY_CURRENT_USER is not implemented."
            Exit Sub
        End If
    
        If Valuetype = REG_SZ Then
            SaveSetting "IguanaTex", Keyname, Valuename, "sz:" & value
        ElseIf Valuetype = REG_DWORD Then
            SaveSetting "IguanaTex", Keyname, Valuename, "dword:" & CStr(value)
        Else
            MsgBox "Error saving registry key."
        End If
    End Sub

#Else

    #If VBA7 Then
        Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
        Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias _
        "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, _
        ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, _
        phkResult As Long, lpdwDisposition As Long) As Long
        Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias _
        "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
        Declare PtrSafe Function RegQueryValueExString Lib "advapi32.dll" Alias _
        "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As LongPtr, _
        ByVal lpReserved As Long, lpType As Long, ByVal lpData As LongPtr, _
        lpcbData As Long) As Long
        Declare PtrSafe Function RegQueryValueExLong Lib "advapi32.dll" Alias _
        "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As LongPtr, _
        ByVal lpReserved As Long, lpType As Long, lpData As Long, _
        lpcbData As Long) As Long
        Declare PtrSafe Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
        "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As LongPtr, _
        ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, _
        lpcbData As Long) As Long
        Declare PtrSafe Function RegSetValueExString Lib "advapi32.dll" Alias _
        "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As LongPtr, _
        ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As LongPtr, _
        ByVal cbData As Long) As Long
        Declare PtrSafe Function RegSetValueExLong Lib "advapi32.dll" Alias _
        "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As LongPtr, _
        ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
        ByVal cbData As Long) As Long
    #Else
        Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
        Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
        "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, _
        ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, _
        phkResult As Long, lpdwDisposition As Long) As Long
        Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
        "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
        Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, _
        lpcbData As Long) As Long
        Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, lpData As Long, _
        lpcbData As Long) As Long
        Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, _
        lpcbData As Long) As Long
        Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, _
        ByVal cbData As Long) As Long
        Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
        ByVal cbData As Long) As Long
    #End If

    Public Function SetValueEx(ByVal hKey As Long, sValueName As String, _
                               lType As Long, vValue As Variant) As Long
        Dim lValue As Long
        Dim sValue As String
        Select Case lType
            Case REG_SZ
                sValue = vValue & Chr$(0)
                #If VBA7 Then
                    SetValueEx = RegSetValueExString(hKey, StrPtr(sValueName), 0&, _
                                               lType, StrPtr(sValue), LenB(sValue))
                #Else
                    SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
                                               lType, sValue, Len(sValue))
                #End If
            Case REG_DWORD
                lValue = vValue
                #If VBA7 Then
                    SetValueEx = RegSetValueExLong(hKey, StrPtr(sValueName), 0&, _
                                                   lType, lValue, 4)
                #Else
                    SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, _
                                                   lType, lValue, 4)
                #End If
            End Select
    End Function

    Public Function QueryValueEx(ByVal lhKey As Long, _
    ByVal szValueName As String, vValue As Variant) As Long
        Dim cch As Long
        Dim lrc As Long
        Dim lType As Long
        Dim lValue As Long
        Dim sValue As String
    
        On Error GoTo QueryValueExError
    
        ' Determine the size and type of data to be read
        #If VBA7 Then
            lrc = RegQueryValueExNULL(lhKey, StrPtr(szValueName), 0&, lType, 0&, cch)
        #Else
            lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
        #End If
        If lrc <> ERROR_NONE Then Err.Raise 5
    
        Select Case lType
            ' For strings
            Case REG_SZ:
                #If VBA7 Then
                    ' Dividing by 2 because cch is in Bytes,
                    ' but String is allocated by number of 2-Byte characters
                    sValue = String(cch / 2, 0)
                    lrc = RegQueryValueExString(lhKey, StrPtr(szValueName), 0&, lType, _
                                                StrPtr(sValue), cch)
                    If lrc = ERROR_NONE Then
                        vValue = Left$(sValue, cch / 2 - 1)
                    Else
                        vValue = Empty
                    End If
                #Else
                    ' For older versions of Office.
                    ' No proper support of Unicode strings, which will be cut
                    sValue = String(cch, 0)
                    lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
                                                sValue, cch)
                    If lrc = ERROR_NONE Then
                        vValue = Left$(sValue, cch - 1)
                    Else
                        vValue = Empty
                    End If
                #End If
            ' For DWORDS
            Case REG_DWORD:
                #If VBA7 Then
                    lrc = RegQueryValueExLong(lhKey, StrPtr(szValueName), 0&, lType, _
                                              lValue, cch)
                #Else
                    lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
                                              lValue, cch)
                #End If

                If lrc = ERROR_NONE Then vValue = lValue
            Case Else
                'all other data types not supported
                lrc = -1
        End Select
    
QueryValueExExit:
        QueryValueEx = lrc
        Exit Function
    
QueryValueExError:
        Resume QueryValueExExit
    End Function

    Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
        Dim hNewKey As Long         'handle to the new key
        Dim lRetVal As Long         'result of the RegCreateKeyEx function
    
        lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
                  vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                  0&, hNewKey, lRetVal)
        RegCloseKey (hNewKey)
    End Sub

    Public Function GetRegistryValue(Hive As Long, Keyname As String, Valuename As String, DefaultValue As Variant) As Variant
        Dim lRetVal As Long         'result of the API functions
        Dim hKey As Long         'handle of opened key
        Dim vValue As Variant      'setting of queried value
        
        lRetVal = RegOpenKeyEx(Hive, Keyname, 0, KEY_QUERY_VALUE, hKey)
        lRetVal = QueryValueEx(hKey, Valuename, vValue)
        RegCloseKey (hKey)
        
        If (lRetVal = 0) Then
            GetRegistryValue = vValue
        Else
            GetRegistryValue = DefaultValue
        End If
    End Function

    Public Sub SetRegistryValue(Hive As Long, ByRef Keyname As String, ByRef Valuename As String, _
Valuetype As Long, value As Variant)
        Dim lRetVal As Long         'result of the SetValueEx function
        Dim hKey As Long         'handle of open key
        
        'open the specified key
        lRetVal = RegOpenKeyEx(Hive, Keyname, 0, KEY_SET_VALUE, hKey)
        If (lRetVal = 0) Then
            lRetVal = SetValueEx(hKey, Valuename, Valuetype, value)
            RegCloseKey (hKey)
        Else
            RegCloseKey (hKey)
            Dim MyKeyname As String
            MyKeyname = Keyname
            Dim MyPredefKey As Long
            MyPredefKey = Hive
            CreateNewKey MyKeyname, MyPredefKey
            lRetVal = RegOpenKeyEx(Hive, Keyname, 0, KEY_SET_VALUE, hKey)
            lRetVal = SetValueEx(hKey, Valuename, Valuetype, value)
            RegCloseKey (hKey)
        
        End If
            
        If (lRetVal <> 0) Then
            MsgBox "Error saving registry key."
        End If
    End Sub

#End If


