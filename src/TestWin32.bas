Attribute VB_Name = "TestWin32"
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

'@TestMethod("win32")
'@IgnoreTest
Private Sub TestOpenNewBrowserInstance()
    On Error GoTo TestFail
    
    'Arrange:
    Dim driver As Long
    
    'Act:
    driver = win32.GetEdgeWindowHandle("http://localhost:8080/")
    
    'Assert:
    Assert.AreNotEqual CLng(0), driver

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    win32.CloseWindow driver
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
