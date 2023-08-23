Attribute VB_Name = "TestTcpServer"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Type TServer
    TcpServer As TcpServer
End Type

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
Private this As TServer

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    Set this.TcpServer = New TcpServer
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set this.TcpServer = Nothing
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
End Sub

'@TestCleanup
Private Sub TestCleanup()
End Sub

'@TestMethod("TcpServer")
Private Sub TestInit()
    On Error GoTo TestFail
    
    'Arrange:
    'We new up the TcpServer on the Module Test
    'We should be able to create a socket with no error if the TcpServer is initialized properly
    
    'Act:
    Dim result As Long
    result = winsock2.socket(winsock2.AF_INET, winsock2.SOCK_STREAM, winsock2.IPPROTO_TCP)
    
    'Assert:
    Assert.AreNotEqual winsock2.INVALID_SOCKET, result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    winsock2.closesocket result
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TcpServer")
Private Sub TestBindAndListenOn()
    On Error GoTo TestFail
    
    'Arrange:
    'We new up the TcpServer on the Module Test
    'We should be able to Bind and Listen on a port
    
    'Act:
    Dim result As Long
    result = this.TcpServer.BindAndListenOn(port:=8080)
    
    'Assert:
    Assert.AreNotEqual winsock2.INVALID_SOCKET, result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

