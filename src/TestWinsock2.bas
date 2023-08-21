Attribute VB_Name = "TestWinsock2"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("winsock2")
Private Sub TestWSAStartup()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsa As winsock2.WSADATA
    
    'Act:
    Dim result As Long
    result = winsock2.WSAStartup(257, wsa)
    
    'Assert:
    Assert.AreEqual CLng(0), result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    winsock2.WSACleanup
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("winsock2")
Private Sub TestWSACleanup()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsa As winsock2.WSADATA
    winsock2.WSAStartup 257, wsa
    
    'Act:
    Dim result As Long
    result = winsock2.WSACleanup
    
    'Assert:
    Assert.AreEqual CLng(0), result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("winsock2")
Private Sub TestSocket()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsa As winsock2.WSADATA
    winsock2.WSAStartup 257, wsa
    
    'Act:
    Dim result As Long
    result = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
    
    'Assert:
    Assert.IsTrue result <> winsock2.INVALID_SOCKET

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    winsock2.closesocket result
    winsock2.WSACleanup
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("winsock2")
Private Sub TestCloseSocket()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsocket As Long
    Dim wsa As winsock2.WSADATA
    winsock2.WSAStartup 257, wsa
    wsocket = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
    
    'Act:
    Dim result As Long
    result = winsock2.closesocket(wsocket)
    
    'Assert:
    Assert.IsTrue result <> winsock2.INVALID_SOCKET

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    winsock2.WSACleanup
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("winsock2")
Private Sub TestHostTONetworkShort_htons()
    On Error GoTo TestFail
    
    'Arrange:
    Dim port As Long
    port = 8080
    
    'Act:
    Dim result As Integer
    result = winsock2.htons(port)
    
    'Assert:
    Assert.AreEqual -28641, result
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("winsock2")
Private Sub TestNetworkTOHostShort_ntohs()
    On Error GoTo TestFail
    
    'Arrange:
    Dim port As Long, netShort As Integer
    port = 8080
    netShort = winsock2.htons(port)
    
    'Act:
    Dim result As Long
    result = winsock2.ntohs(netShort)
    
    'Assert:
    Assert.AreEqual port, result
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("winsock2")
Private Sub TestBind()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsa As winsock2.WSADATA
    winsock2.WSAStartup 257, wsa
    
    Dim socket As Long
    socket = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
    
    Dim port As Long
    port = 8080
    
    Dim endpoint As winsock2.sockaddr_in
    endpoint.sin_family = winsock2.AF_INET
    endpoint.sin_addr.s_addr = winsock2.INADDR_ANY
    endpoint.sin_port = winsock2.htons(port)
    
    'Act:
    Dim result As Long
    result = winsock2.bind(socket, endpoint, 16)
    
    'Assert:
    Assert.AreNotEqual winsock2.SOCKET_ERROR, result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    winsock2.closesocket socket
    winsock2.WSACleanup
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("winsock2")
Private Sub TestListen()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsa As winsock2.WSADATA
    winsock2.WSAStartup 257, wsa
    
    Dim socket As Long
    socket = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
    
    Dim port As Long
    port = 8080
    
    Dim endpoint As winsock2.sockaddr_in
    endpoint.sin_family = winsock2.AF_INET
    endpoint.sin_addr.s_addr = winsock2.INADDR_ANY
    endpoint.sin_port = winsock2.htons(port)
    
    winsock2.bind socket, endpoint, 16
    
    'Act:
    Dim result As Long
    result = winsock2.listen(socket, 10)
    
    'Assert:
    Assert.AreNotEqual winsock2.SOCKET_ERROR, result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    winsock2.closesocket socket
    winsock2.WSACleanup
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

''@TestMethod("winsock2")
'Private Sub TestSelect_returnsNumberOfSocketHandles()
'    On Error GoTo TestFail
'    'The select function returns the total number of socket handles that are ready and contained in the fd_set structures
'    'Arrange:
'    Dim wsa As winsock2.WSADATA
'    winsock2.WSAStartup 257, wsa
'
'    Dim socket As Long
'    socket = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
'
'    Dim port As Long
'    port = 8080
'
'    Dim endpoint As winsock2.sockaddr_in
'    endpoint.sin_family = winsock2.AF_INET
'    endpoint.sin_addr.s_addr = winsock2.INADDR_ANY
'    endpoint.sin_port = winsock2.htons(port)
'
'    winsock2.bind socket, endpoint, 16
'
'    winsock2.listen socket, CInt(10)
'
'    Dim readFDS As winsock2.FD_SET, emptyFDS As winsock2.FD_SET
'    winsock2.FD_ZERO readFDS
'
'    Dim time As winsock2.timeval
'    time.tv_sec = 10
'    time.tv_usec = 0
'
'    Dim driver As Long
'    driver = win32.GetEdgeWindowHandle("http://localhost:8080/")
'
'    winsock2.FD_SET socket, readFDS
'
'    'Act:
'    Dim result As Integer
'    result = winsock2.select_(socket + 1, readFDS, emptyFDS, emptyFDS, time)
'
'    'Assert:
'    Assert.AreEqual 1, result
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    Debug.Print winsock2.WSAGetLastError()
'    On Error Resume Next
'
'    win32.CloseWindow driver
'
'    winsock2.closesocket socket
'    winsock2.WSACleanup
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

'@TestMethod("winsock2")
Private Sub TestAccept()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsa As winsock2.WSADATA
    winsock2.WSAStartup 257, wsa
    
    Dim socket As Long
    socket = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
    
    Dim port As Long
    port = 8080
    
    Dim endpoint As winsock2.sockaddr_in
    endpoint.sin_family = winsock2.AF_INET
    endpoint.sin_addr.s_addr = winsock2.INADDR_ANY
    endpoint.sin_port = winsock2.htons(port)
    
    winsock2.bind socket, endpoint, 16
    winsock2.listen socket, 10
    
    Dim socketAddress As winsock2.sockaddr
    
    Dim driver As Long
    driver = win32.GetEdgeWindowHandle("http://localhost:8080/")
    
    'Act:
    Dim result As Long
    result = winsock2.accept(socket, socketAddress, 16)
    
    'Assert:
    Assert.AreNotEqual winsock2.INVALID_SOCKET, result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    win32.CloseWindow driver
    
    winsock2.closesocket socket
    winsock2.WSACleanup
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("winsock2")
Private Sub TestSetsockopt()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsa As winsock2.WSADATA
    winsock2.WSAStartup 257, wsa
    
    Dim socket As Long
    socket = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
    
    Dim port As Long
    port = 8080
    
    Dim endpoint As winsock2.sockaddr_in
    endpoint.sin_family = winsock2.AF_INET
    endpoint.sin_addr.s_addr = winsock2.INADDR_ANY
    endpoint.sin_port = winsock2.htons(port)
    
    winsock2.bind socket, endpoint, 16
    winsock2.listen socket, 10
    
    Dim socketAddress As winsock2.sockaddr
    
    Dim driver As Long
    driver = win32.GetEdgeWindowHandle("http://localhost:8080/")
    
    Dim clientSocket As Long
    clientSocket = winsock2.accept(socket, socketAddress, 16)
    
    Dim timeout As Long
    timeout = 10
    
    'Act:
    Dim result As Long
    result = winsock2.setsockopt(clientSocket, winsock2.SOL_SOCKET, winsock2.SO_RCVTIMEO, timeout, 4)
    
    'Assert:
    Assert.AreNotEqual winsock2.SOCKET_ERROR, result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    win32.CloseWindow driver
    
    winsock2.closesocket socket
    winsock2.WSACleanup
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
