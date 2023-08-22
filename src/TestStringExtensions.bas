Attribute VB_Name = "TestStringExtensions"
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

'@TestMethod("StringExtensions")
Private Sub TestRepeat()
    On Error GoTo TestFail
    
    'Arrange:
    Const text As String = "a"
    Const count As Long = 3
    
    'Act:
    Dim result As String
    result = StringExtensions.Repeat(text, count)
    
    'Assert:
    Assert.AreEqual "aaa", result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringExtensions")
Private Sub TestSubstring()
    On Error GoTo TestFail
    
    'Arrange:
    Const buffer As String = "abacaba"
    Const startIndex As Long = 4
    Const readBytes As Long = 4
    
    'Act:
    Dim result As String
    result = StringExtensions.Substring(buffer, startIndex, readBytes)
    
    'Assert:
    Assert.AreEqual "caba", result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
