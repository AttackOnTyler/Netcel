Attribute VB_Name = "TextEncoding"
'@Folder("app.resources.stringUtils")
Option Explicit

'Byte <-> String conversion. Raw socket data is held as a "binary string" (Latin-1,
'one char per byte) so that lengths and offsets always match byte counts.

Private Const CP_LATIN1 As Long = 28591
Private Const CP_UTF8 As Long = 65001

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByVal multiByteStr As LongPtr, ByVal multiByteLength As Long, ByVal wideCharStr As LongPtr, ByVal wideCharLength As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByVal wideCharStr As LongPtr, ByVal wideCharLength As Long, ByVal multiByteStr As LongPtr, ByVal multiByteLength As Long, ByVal defaultChar As LongPtr, ByVal usedDefaultChar As LongPtr) As Long

Public Function Utf8Encode(ByVal text As String) As Byte()
    Utf8Encode = Encode(text, CP_UTF8)
End Function

Public Function Utf8Decode(ByRef bytes() As Byte, Optional ByVal count As Long = -1) As String
    Utf8Decode = Decode(bytes, count, CP_UTF8)
End Function

'Bytes -> binary string (each byte becomes one char U+0000..U+00FF).
Public Function BytesToBinaryString(ByRef bytes() As Byte, Optional ByVal count As Long = -1) As String
    BytesToBinaryString = Decode(bytes, count, CP_LATIN1)
End Function

'Binary string -> bytes (inverse of BytesToBinaryString).
Public Function BinaryStringToBytes(ByVal text As String) As Byte()
    BinaryStringToBytes = Encode(text, CP_LATIN1)
End Function

Private Function Encode(ByVal text As String, ByVal codePage As Long) As Byte()
    Dim result() As Byte
    If Len(text) = 0 Then
        result = vbNullString 'empty array, UBound = -1
        Encode = result
        Exit Function
    End If

    Dim size As Long
    size = WideCharToMultiByte(codePage, 0, StrPtr(text), Len(text), 0, 0, 0, 0)
    ReDim result(0 To size - 1)
    WideCharToMultiByte codePage, 0, StrPtr(text), Len(text), VarPtr(result(0)), size, 0, 0
    Encode = result
End Function

Private Function Decode(ByRef bytes() As Byte, ByVal count As Long, ByVal codePage As Long) As String
    If count < 0 Then count = ByteCount(bytes)
    If count <= 0 Then Exit Function

    Dim size As Long
    size = MultiByteToWideChar(codePage, 0, VarPtr(bytes(LBound(bytes))), count, 0, 0)
    Decode = String$(size, vbNullChar)
    MultiByteToWideChar codePage, 0, VarPtr(bytes(LBound(bytes))), count, StrPtr(Decode), size
End Function

Public Function ByteCount(ByRef bytes() As Byte) As Long
    On Error GoTo Empty_
    ByteCount = UBound(bytes) - LBound(bytes) + 1
    Exit Function
Empty_:
    ByteCount = 0
End Function
