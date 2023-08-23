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

Private Type TTcpServer
    socket As Long
End Type

Private this As TTcpServer

Private Sub Class_Initialize()
    Dim wsa As winsock2.WSADATA
    winsock2.WSAStartup 257, wsa
End Sub

Public Function BindAndListenOn(ByVal port As Long, Optional ByVal backlog As Integer = 10) As Long
    this.socket = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
    
    If this.socket = winsock2.INVALID_SOCKET Then GoTo SocketFail
    
    Dim endpoint As winsock2.sockaddr_in
    endpoint.sin_family = winsock2.AF_INET
    endpoint.sin_addr.s_addr = winsock2.INADDR_ANY
    endpoint.sin_port = winsock2.htons(port)
    
    If winsock2.bind(this.socket, endpoint, 16) = winsock2.SOCKET_ERROR Then GoTo BindAndListenOnFail
    
    If winsock2.listen(this.socket, backlog) = winsock2.SOCKET_ERROR Then GoTo BindAndListenOnFail
    
BindAndListenOnExit:
    Exit Function
BindAndListenOnFail:
    winsock2.closesocket this.socket
SocketFail:
    BindAndListenOn = -1
    Resume BindAndListenOnExit
End Function

Private Function endpoint(ByVal port As Long) As winsock2.sockaddr_in
    Dim sockaddr_in As winsock2.sockaddr_in
    sockaddr_in.sin_family = winsock2.AF_INET
    sockaddr_in.sin_addr.s_addr = winsock2.INADDR_ANY
    sockaddr_in.sin_port = winsock2.htons(port)
    endpoint = sockaddr_in
End Function

Private Sub Class_Terminate()
    If this.socket <> 0 Then winsock2.closesocket this.socket
    winsock2.WSACleanup
End Sub
