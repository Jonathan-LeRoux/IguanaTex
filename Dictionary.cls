VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Dictionary v1.4.1
' (c) Tim Hall - https://github.com/timhall/VBA-Dictionary
'
' Drop-in replacement for Scripting.Dictionary on Mac
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

#Const UseScriptingDictionaryIfAvailable = True

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

' dict_KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
Private dict_pKeyValues As collection
Private dict_pKeys() As Variant
Private dict_pItems() As Variant
Private dict_pObjectKeys As collection
Private dict_pCompareMode As CompareMethod

#Else

Private dict_pDictionary As Object

#End If

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum CompareMethod
    BinaryCompare = VBA.vbBinaryCompare
    TextCompare = VBA.vbTextCompare
    DatabaseCompare = VBA.vbDatabaseCompare
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Property Get CompareMode() As CompareMethod
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    CompareMode = dict_pCompareMode
#Else
    CompareMode = dict_pDictionary.CompareMode
#End If
End Property
Public Property Let CompareMode(value As CompareMethod)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.count > 0 Then
        ' Can't change CompareMode for Dictionary that contains data
        ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
        Err.Raise 5 ' Invalid procedure call or argument
    End If

    dict_pCompareMode = value
#Else
    dict_pDictionary.CompareMode = value
#End If
End Property

Public Property Get count() As Long
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    count = dict_pKeyValues.count
#Else
    count = dict_pDictionary.count
#End If
End Property

Public Property Get item(Key As Variant) As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Key)

    If Not IsEmpty(dict_KeyValue) Then
        If VBA.IsObject(dict_KeyValue(2)) Then
            Set item = dict_KeyValue(2)
        Else
            item = dict_KeyValue(2)
        End If
    Else
        ' Not found -> Returns Empty
    End If
#Else
    If VBA.IsObject(dict_pDictionary.item(Key)) Then
        Set item = dict_pDictionary.item(Key)
    Else
        item = dict_pDictionary.item(Key)
    End If
#End If
End Property
Public Property Let item(Key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(Key) Then
        dict_ReplaceKeyValue dict_GetKeyValue(Key), Key, value
    Else
        dict_AddKeyValue Key, value
    End If
#Else
    dict_pDictionary.item(Key) = value
#End If
End Property
Public Property Set item(Key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(Key) Then
        dict_ReplaceKeyValue dict_GetKeyValue(Key), Key, value
    Else
        dict_AddKeyValue Key, value
    End If
#Else
    Set dict_pDictionary.item(Key) = value
#End If
End Property

Public Property Let Key(Previous As Variant, Updated As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Previous)

    If Not VBA.IsEmpty(dict_KeyValue) Then
        dict_ReplaceKeyValue dict_KeyValue, Updated, dict_KeyValue(2)
    End If
#Else
    dict_pDictionary.Key(Previous) = Updated
#End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add an item with the given key
'
' @param {Variant} Key
' @param {Variant} Item
' --------------------------------------------- '
Public Sub Add(Key As Variant, item As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Not Me.Exists(Key) Then
        dict_AddKeyValue Key, item
    Else
        ' This key is already associated with an element of this collection
        Err.Raise 457
    End If
#Else
    dict_pDictionary.Add Key, item
#End If
End Sub

''
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(Key As Variant) As Boolean
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Exists = Not IsEmpty(dict_GetKeyValue(Key))
#Else
    Exists = dict_pDictionary.Exists(Key)
#End If
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function items() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.count > 0 Then
        items = dict_pItems
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        items = VBA.Split("")
    End If
#Else
    items = dict_pDictionary.items
#End If
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.count > 0 Then
        Keys = dict_pKeys
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        Keys = VBA.Split("")
    End If
#Else
    Keys = dict_pDictionary.Keys
#End If
End Function

''
' Remove an item for the given key
'
' @param {Variant} Key
' --------------------------------------------- '
Public Sub Remove(Key As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Key)

    If Not VBA.IsEmpty(dict_KeyValue) Then
        dict_RemoveKeyValue dict_KeyValue
    Else
        ' Application-defined or object-defined error
        Err.Raise 32811
    End If
#Else
    dict_pDictionary.Remove Key
#End If
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set dict_pKeyValues = New collection

    Erase dict_pKeys
    Erase dict_pItems
#Else
    dict_pDictionary.RemoveAll
#End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

Private Function dict_GetKeyValue(dict_Key As Variant) As Variant
    On Error Resume Next
    dict_GetKeyValue = dict_pKeyValues(dict_GetFormattedKey(dict_Key))
    Err.Clear
End Function

Private Sub dict_AddKeyValue(dict_Key As Variant, dict_Value As Variant, Optional dict_Index As Long = -1)
    If Me.count = 0 Then
        ReDim dict_pKeys(0 To 0)
        ReDim dict_pItems(0 To 0)
    Else
        ReDim Preserve dict_pKeys(0 To UBound(dict_pKeys) + 1)
        ReDim Preserve dict_pItems(0 To UBound(dict_pItems) + 1)
    End If

    Dim dict_FormattedKey As String
    dict_FormattedKey = dict_GetFormattedKey(dict_Key)

    If dict_Index >= 0 And dict_Index < dict_pKeyValues.count Then
        ' Shift keys/items after + including index into empty last slot
        Dim dict_i As Long
        For dict_i = UBound(dict_pKeys) To dict_Index + 1 Step -1
            dict_pKeys(dict_i) = dict_pKeys(dict_i - 1)
            If VBA.IsObject(dict_pItems(dict_i - 1)) Then
                Set dict_pItems(dict_i) = dict_pItems(dict_i - 1)
            Else
                dict_pItems(dict_i) = dict_pItems(dict_i - 1)
            End If
        Next dict_i

        ' Add key/item at index
        dict_pKeys(dict_Index) = dict_Key
        If VBA.IsObject(dict_Value) Then
            Set dict_pItems(dict_Index) = dict_Value
        Else
            dict_pItems(dict_Index) = dict_Value
        End If

        ' Add key-value at proper index
        dict_pKeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey, before:=dict_Index + 1
    Else
        ' Add key-value as last item
        If VBA.IsObject(dict_Key) Then
            Set dict_pKeys(UBound(dict_pKeys)) = dict_Key
        Else
            dict_pKeys(UBound(dict_pKeys)) = dict_Key
        End If
        If VBA.IsObject(dict_Value) Then
            Set dict_pItems(UBound(dict_pItems)) = dict_Value
        Else
            dict_pItems(UBound(dict_pItems)) = dict_Value
        End If

        dict_pKeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey
    End If
End Sub

Private Sub dict_ReplaceKeyValue(dict_KeyValue As Variant, dict_Key As Variant, dict_Value As Variant)
    Dim dict_Index As Long
    Dim dict_i As Integer

    dict_Index = dict_GetKeyIndex(dict_KeyValue(1))

    ' Remove existing dict_Value
    dict_RemoveKeyValue dict_KeyValue, dict_Index

    ' Add new dict_Key dict_Value back
    dict_AddKeyValue dict_Key, dict_Value, dict_Index
End Sub

Private Sub dict_RemoveKeyValue(dict_KeyValue As Variant, Optional ByVal dict_Index As Long = -1)
    Dim dict_i As Long
    If dict_Index = -1 Then
        dict_Index = dict_GetKeyIndex(dict_KeyValue(1))
    End If

    If dict_Index >= 0 And dict_Index <= UBound(dict_pKeys) Then
        ' Shift keys/items after index down
        For dict_i = dict_Index To UBound(dict_pKeys) - 1
            dict_pKeys(dict_i) = dict_pKeys(dict_i + 1)

            If VBA.IsObject(dict_pItems(dict_i + 1)) Then
                Set dict_pItems(dict_i) = dict_pItems(dict_i + 1)
            Else
                dict_pItems(dict_i) = dict_pItems(dict_i + 1)
            End If
        Next dict_i

        ' Resize keys/items to remove empty slot
        If UBound(dict_pKeys) = 0 Then
            Erase dict_pKeys
            Erase dict_pItems
        Else
            ReDim Preserve dict_pKeys(0 To UBound(dict_pKeys) - 1)
            ReDim Preserve dict_pItems(0 To UBound(dict_pItems) - 1)
        End If
    End If

    dict_pKeyValues.Remove dict_KeyValue(0)
    dict_RemoveObjectKey dict_KeyValue(1)
End Sub

Private Function dict_GetFormattedKey(dict_Key As Variant) As String
    If VBA.IsObject(dict_Key) Then
        dict_GetFormattedKey = dict_GetObjectKey(dict_Key)
    ElseIf VarType(dict_Key) = VBA.vbBoolean Then
        dict_GetFormattedKey = IIf(dict_Key, "-1__-1", "0__0")
    ElseIf VarType(dict_Key) = VBA.vbString Then
        dict_GetFormattedKey = dict_Key

        If Me.CompareMode = CompareMethod.BinaryCompare Then
            ' Collection does not have method of setting key comparison
            ' So case-sensitive keys aren't supported by default
            ' -> Approach: Append lowercase characters to original key
            '    AbC -> AbC___b_, abc -> abc__abc, ABC -> ABC_____
            Dim dict_Lowercase As String
            dict_Lowercase = ""

            Dim dict_i As Integer
            Dim dict_Char As String
            Dim dict_Ascii As Integer
            For dict_i = 1 To VBA.Len(dict_GetFormattedKey)
                dict_Char = VBA.Mid$(dict_GetFormattedKey, dict_i, 1)
                dict_Ascii = VBA.Asc(dict_Char)
                If dict_Ascii >= 97 And dict_Ascii <= 122 Then
                    dict_Lowercase = dict_Lowercase & dict_Char
                Else
                    dict_Lowercase = dict_Lowercase & "_"
                End If
            Next dict_i

            If dict_Lowercase <> "" Then
                dict_GetFormattedKey = dict_GetFormattedKey & "__" & dict_Lowercase
            End If
        End If
    Else
        ' For numbers, add duplicate to distinguish from strings
        ' ->  123  -> "123__123"
        '    "123" -> "123"
        dict_GetFormattedKey = VBA.CStr(dict_Key) & "__" & CStr(dict_Key)
    End If
End Function

Private Function dict_GetObjectKey(dict_ObjKey As Variant) As String
    Dim dict_i As Integer
    For dict_i = 1 To dict_pObjectKeys.count
        If dict_pObjectKeys.item(dict_i) Is dict_ObjKey Then
            dict_GetObjectKey = "__object__" & dict_i
            Exit Function
        End If
    Next dict_i

    dict_pObjectKeys.Add dict_ObjKey
    dict_GetObjectKey = "__object__" & dict_pObjectKeys.count
End Function

Private Sub dict_RemoveObjectKey(dict_ObjKey As Variant)
    Dim dict_i As Integer
    For dict_i = 1 To dict_pObjectKeys.count
        If dict_pObjectKeys.item(dict_i) Is dict_ObjKey Then
            dict_pObjectKeys.Remove dict_i
            Exit Sub
        End If
    Next dict_i
End Sub

Private Function dict_GetKeyIndex(dict_Key As Variant) As Long
    Dim dict_i As Long
    For dict_i = 0 To UBound(dict_pKeys)
        If VBA.IsObject(dict_pKeys(dict_i)) And VBA.IsObject(dict_Key) Then
            If dict_pKeys(dict_i) Is dict_Key Then
                dict_GetKeyIndex = dict_i
                Exit For
            End If
        ElseIf VBA.IsObject(dict_pKeys(dict_i)) Or VBA.IsObject(dict_Key) Then
            ' Both need to be objects to check equality, skip
        ElseIf dict_pKeys(dict_i) = dict_Key Then
            dict_GetKeyIndex = dict_i
            Exit For
        End If
    Next dict_i
End Function

#End If

Private Sub Class_Initialize()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set dict_pKeyValues = New collection

    Erase dict_pKeys
    Erase dict_pItems
    Set dict_pObjectKeys = New collection
#Else
    Set dict_pDictionary = CreateObject("Scripting.Dictionary")
#End If
End Sub

Private Sub Class_Terminate()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set dict_pKeyValues = Nothing
    Set dict_pObjectKeys = Nothing
#Else
    Set dict_pDictionary = Nothing
#End If
End Sub


