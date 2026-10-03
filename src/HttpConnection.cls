VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("app.http")
Option Explicit

'One accepted client socket. Service() is called repeatedly from the server tick; it
'reads whatever is available without blocking, and once a full request (head + body
'per Content-Length) has arrived it dispatches it, writes the response and closes.
'One request per connection (Connection: close).

Private Const RECV_CHUNK As Long = 4096
Private Const MAX_HEAD_BYTES As Long = 16384
Private Const MAX_BODY_BYTES As Long = 1048576
Private Const REQUEST_TIMEOUT_SECONDS As Double = 10
Private Const SEND_TIMEOUT_SECONDS As Double = 2

Private Type THttpConnection
    socket As LongPtr
    raw As String 'binary string: one char per received byte
    startedAt As Double
    closed As Boolean
End Type

Private this As THttpConnection

Private Sub Class_Initialize()
    this.socket = winsock2.INVALID_SOCKET
    this.closed = True
End Sub

Public Sub Init(ByVal client As LongPtr)
    this.socket = client
    this.raw = vbNullString
    this.startedAt = Timer
    this.closed = (client = winsock2.INVALID_SOCKET)
End Sub

Public Property Get IsClosed() As Boolean
    IsClosed = this.closed
End Property

'Advances the connection. Returns True once the connection is finished and closed.
Public Function Service(ByVal appRouter As Router) As Boolean
    If this.closed Then
        Service = True
        Exit Function
    End If

    Dim peerClosed As Boolean
    If Not ReadAvailable(peerClosed) Then
        Abort
        Service = True
        Exit Function
    End If

    Dim headEnd As Long
    headEnd = InStr(1, this.raw, vbCrLf & vbCrLf, vbBinaryCompare)

    If headEnd = 0 Then
        If Len(this.raw) > MAX_HEAD_BYTES Then
            Respond Http.NewResponse(431, "request head too large" & vbLf)
        ElseIf peerClosed Then
            Abort
        ElseIf TimedOut() Then
            Respond Http.NewResponse(408, "request timeout" & vbLf)
        Else
            Exit Function 'wait for more data
        End If
        Service = True
        Exit Function
    End If

    If headEnd - 1 > MAX_HEAD_BYTES Then
        Respond Http.NewResponse(431, "request head too large" & vbLf)
        Service = True
        Exit Function
    End If

    Dim request As HttpRequest
    Set request = New HttpRequest
    If Not request.ParseHead(Left$(this.raw, headEnd - 1)) Then
        Respond Http.NewResponse(400, "bad request" & vbLf)
        Service = True
        Exit Function
    End If

    If request.HasHeader("Transfer-Encoding") Then
        Respond Http.NewResponse(501, "chunked request bodies are not supported" & vbLf)
        Service = True
        Exit Function
    End If

    Dim contentLength As Long
    If request.HasHeader("Content-Length") Then
        Dim lengthText As String
        lengthText = request.Header("Content-Length")
        If Not IsDecimalDigits(lengthText) Or Len(lengthText) > 9 Then
            Respond Http.NewResponse(400, "bad Content-Length" & vbLf)
            Service = True
            Exit Function
        End If
        contentLength = CLng(lengthText)
    End If

    If contentLength > MAX_BODY_BYTES Then
        Respond Http.NewResponse(413, "request body too large" & vbLf)
        Service = True
        Exit Function
    End If

    Dim bodyStart As Long
    bodyStart = headEnd + 4
    If Len(this.raw) - bodyStart + 1 < contentLength Then
        If peerClosed Then
            Abort
        ElseIf TimedOut() Then
            Respond Http.NewResponse(408, "request timeout" & vbLf)
        Else
            Exit Function 'wait for the rest of the body
        End If
        Service = True
        Exit Function
    End If

    If contentLength > 0 Then
        Dim bodyBytes() As Byte
        bodyBytes = TextEncoding.BinaryStringToBytes(Mid$(this.raw, bodyStart, contentLength))
        request.Body = TextEncoding.Utf8Decode(bodyBytes)
    End If

    Dim response As HttpResponse
    Set response = Handle(appRouter, request)
    Debug.Print Format$(Now, "hh:nn:ss"); " "; request.Method; " "; request.Target; " -> "; response.StatusCode
    Respond response
    Service = True
End Function

'Force-closes the connection (used on server shutdown).
Public Sub Abort()
    If this.socket <> winsock2.INVALID_SOCKET Then
        winsock2.closesocket this.socket
        this.socket = winsock2.INVALID_SOCKET
    End If
    this.closed = True
End Sub

Private Function Handle(ByVal appRouter As Router, ByVal request As HttpRequest) As HttpResponse
    On Error GoTo Fail
    Set Handle = appRouter.Dispatch(request)
    If Handle Is Nothing Then Set Handle = Http.NewResponse(500, "handler returned no response" & vbLf)
    Exit Function
Fail:
    Set Handle = Http.NewResponse(500, "handler error " & Err.Number & ": " & Err.Description & vbLf)
End Function

'Reads everything currently available. Returns False on a hard socket error.
Private Function ReadAvailable(ByRef peerClosed As Boolean) As Boolean
    Dim buffer() As Byte
    ReDim buffer(0 To RECV_CHUNK - 1)

    Dim received As Long
    Do
        received = winsock2.recv(this.socket, buffer(0), RECV_CHUNK, 0)
        If received > 0 Then
            this.raw = this.raw & TextEncoding.BytesToBinaryString(buffer, received)
            If Len(this.raw) > MAX_HEAD_BYTES + MAX_BODY_BYTES Then Exit Do
        ElseIf received = 0 Then
            peerClosed = True
            Exit Do
        ElseIf winsock2.WSAGetLastError() = winsock2.WSAEWOULDBLOCK Then
            Exit Do
        Else
            Exit Function
        End If
    Loop

    ReadAvailable = True
End Function

Private Sub Respond(ByVal response As HttpResponse)
    Dim bytes() As Byte
    bytes = response.ToBytes()
    SendAll bytes
    If this.socket <> winsock2.INVALID_SOCKET Then winsock2.shutdown this.socket, winsock2.SD_SEND
    Abort
End Sub

'send() on a non-blocking socket may accept only part of the buffer or report
'WSAEWOULDBLOCK; retry briefly rather than block forever.
Private Sub SendAll(ByRef bytes() As Byte)
    Dim total As Long, offset As Long, sent As Long, started As Double
    total = TextEncoding.ByteCount(bytes)
    started = Timer

    Do While offset < total
        sent = winsock2.send(this.socket, bytes(offset), total - offset, 0)
        If sent = winsock2.SOCKET_ERROR Then
            If winsock2.WSAGetLastError() <> winsock2.WSAEWOULDBLOCK Then Exit Sub
            If SecondsSince(started) > SEND_TIMEOUT_SECONDS Then Exit Sub
            win32.Sleep 1
        Else
            offset = offset + sent
        End If
    Loop
End Sub

Private Function TimedOut() As Boolean
    TimedOut = SecondsSince(this.startedAt) > REQUEST_TIMEOUT_SECONDS
End Function

Private Function SecondsSince(ByVal startedAt As Double) As Double
    SecondsSince = Timer - startedAt
    If SecondsSince < 0 Then SecondsSince = SecondsSince + 86400 'Timer wraps at midnight
End Function

Private Function IsDecimalDigits(ByVal text As String) As Boolean
    If Len(text) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(text)
        If Mid$(text, i, 1) < "0" Or Mid$(text, i, 1) > "9" Then Exit Function
    Next
    IsDecimalDigits = True
End Function

Private Sub Class_Terminate()
    Abort
End Sub
