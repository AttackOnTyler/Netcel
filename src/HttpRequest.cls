VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("app.http")
Option Explicit

Private Type THttpRequest
    method As String
    target As String
    path As String
    query As String
    version As String
    headers As Object 'Scripting.Dictionary, case-insensitive keys
    body As String
End Type

Private this As THttpRequest

Private Sub Class_Initialize()
    Set this.headers = CreateObject("Scripting.Dictionary")
    this.headers.CompareMode = vbTextCompare
End Sub

Public Property Get Method() As String
    Method = this.method
End Property

'Request target as sent, e.g. "/ping?x=1".
Public Property Get Target() As String
    Target = this.target
End Property

'Target without the query string, e.g. "/ping".
Public Property Get Path() As String
    Path = this.path
End Property

Public Property Get Query() As String
    Query = this.query
End Property

Public Property Get Version() As String
    Version = this.version
End Property

Public Property Get Body() As String
    Body = this.body
End Property

Public Property Let Body(ByVal value As String)
    this.body = value
End Property

Public Function HasHeader(ByVal name As String) As Boolean
    HasHeader = this.headers.Exists(name)
End Function

Public Function Header(ByVal name As String) As String
    If this.headers.Exists(name) Then Header = this.headers(name)
End Function

'Parses the request line and header block (everything before the blank line,
'CRLF separated). Returns False if the head is malformed.
Public Function ParseHead(ByVal head As String) As Boolean
    Dim lines() As String
    lines = Split(head, vbCrLf)
    If UBound(lines) < 0 Then Exit Function

    Dim requestLine() As String
    requestLine = Split(lines(0), " ")
    If UBound(requestLine) <> 2 Then Exit Function

    this.method = requestLine(0)
    this.target = requestLine(1)
    this.version = requestLine(2)
    If Len(this.method) = 0 Or Len(this.target) = 0 Then Exit Function
    If Left$(this.version, 5) <> "HTTP/" Then Exit Function

    Dim queryStart As Long
    queryStart = InStr(1, this.target, "?")
    If queryStart > 0 Then
        this.path = Left$(this.target, queryStart - 1)
        this.query = Mid$(this.target, queryStart + 1)
    Else
        this.path = this.target
    End If

    Dim i As Long, colon As Long, name As String, value As String
    For i = 1 To UBound(lines)
        If Len(lines(i)) > 0 Then
            colon = InStr(1, lines(i), ":")
            If colon <= 1 Then Exit Function
            name = Trim$(Left$(lines(i), colon - 1))
            value = Trim$(Mid$(lines(i), colon + 1))
            If this.headers.Exists(name) Then
                this.headers(name) = this.headers(name) & ", " & value
            Else
                this.headers.Add name, value
            End If
        End If
    Next

    ParseHead = True
End Function
