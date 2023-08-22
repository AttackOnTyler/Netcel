Attribute VB_Name = "winsock2"
'@Folder("app.resources.win32")
Option Explicit

Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYS_STATUS_LEN = 128

Public Const WSADESCRIPTION_LEN_ARRAY = WSADESCRIPTION_LEN + 1
Public Const WSASYS_STATUS_LEN_ARRAY = WSASYS_STATUS_LEN + 1

Public Const AF_INET = 2
Public Const SOCK_STREAM = 1
Public Const IPPROTO_TCP = 6
Public Const INADDR_ANY = 0

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN_ARRAY
    szSystemStatus As String * WSASYS_STATUS_LEN_ARRAY
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type

Public Type IN_ADDR
    s_addr As Long
End Type

Public Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As IN_ADDR
    sin_zero As String * 8
End Type

'Public Const FD_SETSIZE = 64
'
'Public Type FD_SET
'    fd_count As Integer
'    fd_array(FD_SETSIZE) As Long
'End Type
'
'Public Type timeval
'    tv_sec As Long
'    tv_usec As Long
'End Type

Public Type sockaddr
    sa_family As Integer
    sa_data As String * 14
End Type

Public Const INVALID_SOCKET As Long = -1
Public Const SOCKET_ERROR As Long = -1

Public Const SOL_SOCKET = 65535
Public Const SO_RCVTIMEO = &H1006

Public Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal versionRequired As Long, wsa As WSADATA) As Long
Public Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
Public Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare PtrSafe Function socket Lib "ws2_32.dll" (ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As Long
Public Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Public Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostShort As Long) As Integer
Public Declare PtrSafe Function ntohs Lib "ws2_32.dll" (ByVal netShort As Integer) As Long
Public Declare PtrSafe Function bind Lib "ws2_32.dll" (ByVal socket As Long, name As sockaddr_in, ByVal nameLength As Integer) As Long
Public Declare PtrSafe Function listen Lib "ws2_32.dll" (ByVal socket As Long, ByVal backlog As Integer) As Long
'Public Declare PtrSafe Function select_ Lib "ws2_32.dll" Alias "select" (ByVal nfds As Integer, readFDS As FD_SET, writefds As FD_SET, exceptfds As FD_SET, timeout As timeval) As Integer
Public Declare PtrSafe Function accept Lib "ws2_32.dll" (ByVal socket As Long, clientAddress As sockaddr, clientAddressLength As Integer) As Long
Public Declare PtrSafe Function setsockopt Lib "ws2_32.dll" (ByVal socket As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Integer) As Long
Public Declare PtrSafe Function recv Lib "ws2_32.dll" (ByVal socket As Long, ByVal buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long
Public Declare PtrSafe Function send Lib "ws2_32.dll" (ByVal socket As Long, buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long

'Public Sub FD_ZERO(ByRef s As FD_SET)
'    s.fd_count = 0
'End Sub
'
'Public Sub FD_SET(ByVal fd As Long, ByRef s As FD_SET)
'    Dim i As Integer
'    i = 0
'
'    Do While i < s.fd_count
'        If s.fd_array(i) = fd Then
'            Exit Do
'        End If
'
'        i = i + 1
'    Loop
'
'    If i = s.fd_count Then
'        If s.fd_count < FD_SETSIZE Then
'            s.fd_array(i) = fd
'            s.fd_count = s.fd_count + 1
'        End If
'    End If
'End Sub
