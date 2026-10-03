Attribute VB_Name = "winsock2"
'@Folder("app.resources.win32")
Option Explicit

'Winsock2 declarations for 64-bit (and 32-bit) VBA7.
'SOCKET is UINT_PTR in C, so every socket handle is a LongPtr here.
'int / u_long parameters are 32-bit (Long); u_short is 16-bit (Integer).

Public Const AF_INET As Long = 2
Public Const SOCK_STREAM As Long = 1
Public Const IPPROTO_TCP As Long = 6
Public Const INADDR_ANY As Long = 0
Public Const INADDR_NONE As Long = -1

Public Const INVALID_SOCKET As Long = -1 'compared against LongPtr; (UINT_PTR)~0 == -1 signed
Public Const SOCKET_ERROR As Long = -1

Public Const SOL_SOCKET As Long = &HFFFF&
Public Const SO_RCVTIMEO As Long = &H1006&
Public Const SO_EXCLUSIVEADDRUSE As Long = -5 '~SO_REUSEADDR

Public Const FIONBIO As Long = &H8004667E 'set/clear non-blocking i/o

Public Const SD_RECEIVE As Long = 0
Public Const SD_SEND As Long = 1
Public Const SD_BOTH As Long = 2

Public Const WSAEWOULDBLOCK As Long = 10035
Public Const WSAEADDRINUSE As Long = 10048

'WSADATA layout differs between 32-bit and 64-bit Windows (field order changes),
'and nothing here reads it, so it is treated as an opaque, generously sized blob.
Public Type WSADATA
    data(0 To 511) As Byte
End Type

Public Type IN_ADDR
    s_addr As Long
End Type

Public Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As IN_ADDR
    sin_zero(0 To 7) As Byte
End Type

Public Const SOCKADDR_IN_SIZE As Long = 16

Public Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal versionRequired As Integer, ByRef wsa As WSADATA) As Long
Public Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long

Public Declare PtrSafe Function socket Lib "ws2_32.dll" (ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As LongPtr
Public Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As LongPtr) As Long
Public Declare PtrSafe Function shutdown Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal how As Long) As Long
Public Declare PtrSafe Function bind Lib "ws2_32.dll" (ByVal s As LongPtr, ByRef name As sockaddr_in, ByVal nameLength As Long) As Long
Public Declare PtrSafe Function listen Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal backlog As Long) As Long
'Pass 0, 0 for addr/addrlen when the peer address is not needed.
Public Declare PtrSafe Function accept Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal addr As LongPtr, ByVal addrLength As LongPtr) As LongPtr
'Pass the first element of a Byte array, e.g. recv(s, buffer(0), size, 0).
Public Declare PtrSafe Function recv Lib "ws2_32.dll" (ByVal s As LongPtr, ByRef buffer As Byte, ByVal bufferLength As Long, ByVal flags As Long) As Long
Public Declare PtrSafe Function send Lib "ws2_32.dll" (ByVal s As LongPtr, ByRef buffer As Byte, ByVal bufferLength As Long, ByVal flags As Long) As Long
Public Declare PtrSafe Function setsockopt Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
Public Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal cmd As Long, ByRef argp As Long) As Long

Public Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostShort As Integer) As Integer
Public Declare PtrSafe Function ntohs Lib "ws2_32.dll" (ByVal netShort As Integer) As Integer
Public Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

'htons for the full 0..65535 port range without Integer overflow in VBA.
Public Function PortToNetwork(ByVal port As Long) As Integer
    Dim swapped As Long
    swapped = ((port And &HFF&) * &H100&) Or ((port \ &H100&) And &HFF&)
    If swapped >= &H8000& Then swapped = swapped - &H10000
    PortToNetwork = CInt(swapped)
End Function

'Puts a socket into (or out of) non-blocking mode. Returns True on success.
Public Function SetNonBlocking(ByVal s As LongPtr, Optional ByVal enabled As Boolean = True) As Boolean
    Dim mode As Long
    If enabled Then mode = 1
    SetNonBlocking = (ioctlsocket(s, FIONBIO, mode) <> SOCKET_ERROR)
End Function
