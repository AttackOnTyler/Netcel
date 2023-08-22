Attribute VB_Name = "StringExtensions"
'@Folder("app.resources.stringUtils")
Option Explicit

Public Function Repeat(ByVal text As String, ByVal count As Long) As String
    If count < 0 Or text = vbNullString Then Exit Function
    Repeat = String(count, text)
End Function

Public Function Substring(ByVal text As String, Optional ByVal startIndex As Long = 1, Optional ByVal length As Long = 1) As String
    If text = vbNullString Or startIndex < 1 Or length < 1 Then Exit Function
    Substring = Mid(text, startIndex, length)
End Function
