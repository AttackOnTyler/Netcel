VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("controllers")
Option Explicit

'A non-blocking listening TCP socket. Winsock is started when the instance is
'created and cleaned up when it is destroyed (WSAStartup/WSACleanup are ref-counted).

Private Type TTcpServer
    socket As LongPtr
    wsaStarted As Boolean
    lastError As Long
End Type

Private this As TTcpServer

Private Sub Class_Initialize()
    Dim wsa As winsock2.WSADATA
    this.socket = winsock2.INVALID_SOCKET
    this.wsaStarted = (winsock2.WSAStartup(&H202, wsa) = 0)
End Sub

Public Property Get Handle() As LongPtr
    Handle = this.socket
End Property

Public Property Get IsListening() As Boolean
    IsListening = (this.socket <> winsock2.INVALID_SOCKET)
End Property

Public Property Get LastError() As Long
    LastError = this.lastError
End Property

'Binds and listens on bindAddress:port in non-blocking mode.
'Returns 0 on success, -1 on failure (see LastError for the WSA error code).
'Defaults to loopback only: this server will eventually expose workbook control.
Public Function BindAndListenOn(ByVal port As Long, Optional ByVal backlog As Long = 10, Optional ByVal bindAddress As String = "127.0.0.1") As Long
    this.lastError = 0
    If Not this.wsaStarted Then
        this.lastError = winsock2.WSAGetLastError()
        BindAndListenOn = -1
        Exit Function
    End If
    If IsListening Then CloseSocket

    this.socket = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
    If this.socket = winsock2.INVALID_SOCKET Then GoTo Fail

    Dim exclusive As Long
    exclusive = 1
    winsock2.setsockopt this.socket, winsock2.SOL_SOCKET, winsock2.SO_EXCLUSIVEADDRUSE, exclusive, 4

    Dim address As winsock2.sockaddr_in
    If Not TryMakeEndpoint(port, bindAddress, address) Then
        this.lastError = 10049 'WSAEADDRNOTAVAIL
        GoTo Cleanup
    End If

    If winsock2.bind(this.socket, address, winsock2.SOCKADDR_IN_SIZE) = winsock2.SOCKET_ERROR Then GoTo Fail
    If winsock2.listen(this.socket, backlog) = winsock2.SOCKET_ERROR Then GoTo Fail
    If Not winsock2.SetNonBlocking(this.socket) Then GoTo Fail

    BindAndListenOn = 0
    Exit Function

Fail:
    this.lastError = winsock2.WSAGetLastError()
Cleanup:
    CloseSocket
    BindAndListenOn = -1
End Function

'Non-blocking accept. Returns a client socket (already non-blocking) or INVALID_SOCKET
'when no connection is pending.
Public Function AcceptClient() As LongPtr
    AcceptClient = winsock2.INVALID_SOCKET
    If Not IsListening Then Exit Function

    Dim client As LongPtr
    client = winsock2.accept(this.socket, 0, 0)
    If client = winsock2.INVALID_SOCKET Then
        this.lastError = winsock2.WSAGetLastError()
        Exit Function
    End If

    winsock2.SetNonBlocking client
    AcceptClient = client
End Function

Public Sub CloseSocket()
    If this.socket <> winsock2.INVALID_SOCKET Then
        winsock2.closesocket this.socket
        this.socket = winsock2.INVALID_SOCKET
    End If
End Sub

Private Function TryMakeEndpoint(ByVal port As Long, ByVal bindAddress As String, ByRef address As winsock2.sockaddr_in) As Boolean
    address.sin_family = winsock2.AF_INET
    address.sin_port = winsock2.PortToNetwork(port)
    If bindAddress = vbNullString Or bindAddress = "0.0.0.0" Then
        address.sin_addr.s_addr = winsock2.INADDR_ANY
    Else
        address.sin_addr.s_addr = winsock2.inet_addr(bindAddress)
        If address.sin_addr.s_addr = winsock2.INADDR_NONE Then Exit Function
    End If
    TryMakeEndpoint = True
End Function

Private Sub Class_Terminate()
    CloseSocket
    If this.wsaStarted Then
        winsock2.WSACleanup
        this.wsaStarted = False
    End If
End Sub
