Attribute VB_Name = "win32"
'@Folder("app.resources.win32")
Option Explicit

'Private Const GW_HWNDNEXT = 2
'Private Const GW_CHILD = 5
Private Const SW_HIDE = 0
'Private Const SW_MINIMIZE = 6
'Private Const WM_GETTEXTLENGTH = 14
Private Const WM_CLOSE = 16

'Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
'Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwprocessid As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function GetEdgeWindowHandle(ByVal url As String) As Long
    Shell "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe --new-window " & url, vbNormalFocus
    Sleep 1000
    GetEdgeWindowHandle = GetForegroundWindow()
End Function

Public Sub CloseWindow(ByVal hwnd As Long)
    SendMessage hwnd, WM_CLOSE, 0, 0
End Sub
