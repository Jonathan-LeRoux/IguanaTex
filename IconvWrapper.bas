Attribute VB_Name = "IconvWrapper"
Option Explicit

#If Mac Then
Private Declare PtrSafe Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal tocode As String, ByVal fromcode As String) As LongLong
Private Declare PtrSafe Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr) As Integer
Private Declare PtrSafe Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr, ByRef inbuf As LongPtr, ByRef inbytesleft As LongLong, ByRef outbuf As LongPtr, ByRef outbytesleft As LongLong) As LongLong

Private Sub RunIconv(ByVal cd As LongPtr, ByRef inBytes() As Byte, ByRef outBytes() As Byte)
    Dim inbuf As LongPtr
    inbuf = VarPtr(inBytes(0))
    Dim inbytesleft As LongLong
    inbytesleft = ArrayLength(inBytes)

    Dim outbuf As LongPtr
    outbuf = VarPtr(outBytes(0))
    Dim outbytesleft As LongLong
    outbytesleft = ArrayLength(outBytes)

    While inbytesleft > 0
        If outbytesleft = 0 Then
            ReDim Preserve outBytes(UBound(outBytes) + CLng(inbytesleft) * 2)
            outbytesleft = CLng(inbytesleft) * 2
        End If
        If iconv(cd, inbuf, inbytesleft, outbuf, outbytesleft) = -1& Then GoTo Error
    Wend

    ReDim Preserve outBytes(UBound(outBytes) - CLng(outbytesleft))
    Exit Sub

Error:
    MsgBox "iconv failed, return empty string"
    ReDim outBytes(0)
End Sub

Public Function Utf8ToString(Utf8() As Byte) As String
    If ArrayLength(Utf8) = 0 Then
        Utf8ToString = ""
        Exit Function
    End If
    
    Dim utf16() As Byte
    ReDim utf16(UBound(Utf8) * 2 + 1)

    Dim cd As LongLong
    cd = iconv_open("utf-16le", "utf-8")
    If cd = -1& Then GoTo Error

    RunIconv cd, Utf8, utf16

    If iconv_close(cd) = -1 Then GoTo Error


    Utf8ToString = utf16
    Exit Function

Error:
    MsgBox "iconv failed, return empty string"
    Utf8ToString = ""
End Function

Public Function StringToUtf8(Str As String) As Byte()
    Dim Utf8() As Byte
    
    If Len(Str) = 0 Then
        StringToUtf8 = Utf8
        Exit Function
    End If
    
    Dim utf16() As Byte
    utf16 = Str
    ' trim zero bytes
    ReDim Preserve utf16(Len(Str) * 2 - 1)

    ReDim Utf8(UBound(utf16))

    Dim cd As LongLong
    cd = iconv_open("utf-8", "utf-16le")
    If cd = -1& Then GoTo Error

    RunIconv cd, utf16, Utf8

    If iconv_close(cd) = -1 Then GoTo Error
    
    StringToUtf8 = Utf8
    Exit Function

Error:
    MsgBox "iconv failed, return empty string"
    ReDim Utf8(0)
    StringToUtf8 = Utf8
End Function


#End If
