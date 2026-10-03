VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("app.http")
Option Explicit

Private Type THttpResponse
    statusCode As Long
    contentType As String
    body As String
    headers As Collection 'of "Name: value" strings
End Type

Private this As THttpResponse

Private Sub Class_Initialize()
    this.statusCode = 200
    this.contentType = "text/plain; charset=utf-8"
    Set this.headers = New Collection
End Sub

Public Property Get StatusCode() As Long
    StatusCode = this.statusCode
End Property

Public Property Let StatusCode(ByVal value As Long)
    this.statusCode = value
End Property

Public Property Get ContentType() As String
    ContentType = this.contentType
End Property

Public Property Let ContentType(ByVal value As String)
    this.contentType = value
End Property

Public Property Get Body() As String
    Body = this.body
End Property

Public Property Let Body(ByVal value As String)
    this.body = value
End Property

'Extra headers. Content-Type, Content-Length, Connection and Server are always written.
Public Sub AddHeader(ByVal name As String, ByVal value As String)
    this.headers.Add name & ": " & value
End Sub

'Serialises the full response (status line, headers, UTF-8 body) to bytes.
Public Function ToBytes() As Byte()
    Dim bodyBytes() As Byte
    bodyBytes = TextEncoding.Utf8Encode(this.body)

    Dim bodyLength As Long
    bodyLength = TextEncoding.ByteCount(bodyBytes)

    Dim head As String
    head = "HTTP/1.1 " & this.statusCode & " " & Http.ReasonPhrase(this.statusCode) & vbCrLf & _
           "Server: Netcel" & vbCrLf & _
           "Content-Type: " & this.contentType & vbCrLf & _
           "Content-Length: " & bodyLength & vbCrLf & _
           "Connection: close" & vbCrLf

    Dim extra As Variant
    For Each extra In this.headers
        head = head & extra & vbCrLf
    Next
    head = head & vbCrLf

    Dim headBytes() As Byte
    headBytes = TextEncoding.BinaryStringToBytes(head)

    Dim headLength As Long
    headLength = TextEncoding.ByteCount(headBytes)

    Dim result() As Byte
    ReDim result(0 To headLength + bodyLength - 1)

    Dim i As Long
    For i = 0 To headLength - 1
        result(i) = headBytes(i)
    Next
    For i = 0 To bodyLength - 1
        result(headLength + i) = bodyBytes(i)
    Next

    ToBytes = result
End Function
