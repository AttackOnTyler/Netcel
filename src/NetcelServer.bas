Attribute VB_Name = "NetcelServer"
'@Folder("app")
Option Explicit

'HTTP server lifecycle. Nothing here blocks: the listening and client sockets are
'non-blocking, and work happens in ServerTick, which Application.OnTime calls about
'once a second while Excel is idle. Each tick accepts pending connections and services
'them for at most SERVICE_WINDOW_SECONDS, so Excel stays responsive.
'
'    NetcelServer.StartServer 8080      'then: curl http://127.0.0.1:8080/ping
'    NetcelServer.StopServer
'
'ThisWorkbook.Workbook_BeforeClose calls OnWorkbookClosing, which stops the server.

Public Const DEFAULT_PORT As Long = 8080
Private Const TICK_PROC As String = "NetcelServer.ServerTick"
Private Const SERVICE_WINDOW_SECONDS As Double = 0.25
Private Const MAX_CONNECTIONS As Long = 32
Private Const STRAY_TICK_VARIABLE As String = "NETCEL_STRAY_TICK"
Private Const STRAY_TICK_WINDOW_SECONDS As Double = 10

'Process-wide environment variables survive the workbook's VBA project being unloaded.
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" (ByVal name As LongPtr, ByVal value As LongPtr) As Long
Private Declare PtrSafe Function GetEnvironmentVariableW Lib "kernel32" (ByVal name As LongPtr, ByVal buffer As LongPtr, ByVal size As Long) As Long

Private Type TServerState
    running As Boolean
    port As Long
    listener As TcpServer
    appRouter As Router
    connections As Collection
    tickScheduled As Boolean
    nextTick As Date
    closing As Boolean 'set by OnWorkbookClosing; survives only while this project stays loaded
End Type

Private state As TServerState

'Starts listening on bindAddress:port (loopback by default). Returns True on success.
Public Function StartServer(Optional ByVal port As Long = DEFAULT_PORT, Optional ByVal bindAddress As String = "127.0.0.1") As Boolean
    If state.running Then StopServer
    state.closing = False
    ClearStrayTickMarker

    Set state.listener = New TcpServer
    If state.listener.BindAndListenOn(port, 10, bindAddress) <> 0 Then
        Debug.Print "Netcel: could not listen on " & bindAddress & ":" & port & " (WSA error " & state.listener.LastError & ")"
        Set state.listener = Nothing 'closes socket, WSACleanup
        Exit Function
    End If

    Set state.appRouter = BuildRouter()
    Set state.connections = New Collection
    state.port = port
    state.running = True
    ScheduleTick

    Debug.Print "Netcel: listening on http://" & bindAddress & ":" & port & "/"
    StartServer = True
End Function

Public Sub StopServer()
    Dim wasRunning As Boolean
    wasRunning = state.running
    state.running = False
    CancelTick

    If Not state.connections Is Nothing Then
        Dim connection As HttpConnection
        For Each connection In state.connections
            connection.Abort
        Next
        Set state.connections = Nothing
    End If

    If Not state.listener Is Nothing Then
        state.listener.CloseSocket
        Set state.listener = Nothing 'Class_Terminate -> WSACleanup
    End If
    Set state.appRouter = Nothing

    If wasRunning Then Debug.Print "Netcel: stopped"
End Sub

Public Function IsRunning() As Boolean
    IsRunning = state.running
End Function

'Called from ThisWorkbook.Workbook_BeforeClose.
'
'Excel silently ignores an OnTime cancellation made inside BeforeClose when the close
'was started from code (Workbook.Close via VBA or COM); a user clicking the close
'button is fine. The ignored tick would then reopen the workbook to run ServerTick.
'So before stopping, leave a process-wide marker; if a tick fires in a freshly
'reopened copy of this workbook, ServerTick sees the marker and closes it again.
Public Sub OnWorkbookClosing()
    If state.tickScheduled Then
        SetStrayTickMarker ThisWorkbook.FullName & "|" & CStr(CDbl(state.nextTick))
    End If
    StopServer
    state.closing = True
End Sub

Public Function BuildRouter() As Router
    Dim result As Router
    Set result = New Router

    Dim handlers As Routes
    Set handlers = New Routes
    handlers.Register result

    Set BuildRouter = result
End Function

'Application.OnTime target. Must stay Public and must never leave an error unhandled
'(that would raise a modal dialog and stop the polling loop).
Public Sub ServerTick()
    On Error GoTo Fail
    state.tickScheduled = False
    If Not state.running Then
        If Not state.closing Then CloseIfReopenedByStrayTick
        Exit Sub
    End If

    ServiceWindow

Reschedule:
    If state.running Then ScheduleTick
    Exit Sub
Fail:
    Debug.Print "Netcel: tick error " & Err.Number & ": " & Err.Description
    Resume Reschedule
End Sub

'A fresh (reopened) project has state.closing = False; a workbook that stayed open
'after OnWorkbookClosing (e.g. the user cancelled a save prompt) has it True.
Private Sub CloseIfReopenedByStrayTick()
    Dim marker As String
    marker = GetStrayTickMarker()
    If Len(marker) = 0 Then Exit Sub

    Dim separator As Long
    separator = InStrRev(marker, "|")
    If separator = 0 Then Exit Sub
    If StrComp(Left$(marker, separator - 1), ThisWorkbook.FullName, vbTextCompare) <> 0 Then Exit Sub

    Dim scheduledFor As Date
    scheduledFor = CDate(CDbl(Mid$(marker, separator + 1)))
    If Abs(Now - scheduledFor) * 86400# > STRAY_TICK_WINDOW_SECONDS Then
        ClearStrayTickMarker
        Exit Sub
    End If

    ClearStrayTickMarker
    state.closing = True
    ThisWorkbook.Close SaveChanges:=False
End Sub

Private Sub ServiceWindow()
    Dim started As Double
    started = Timer

    Do
        AcceptPending
        ServiceConnections
        If state.connections.Count = 0 Then Exit Do
        If SecondsSince(started) > SERVICE_WINDOW_SECONDS Then Exit Do
        win32.Sleep 5
    Loop
End Sub

Private Sub AcceptPending()
    Dim client As LongPtr, connection As HttpConnection
    Do While state.connections.Count < MAX_CONNECTIONS
        client = state.listener.AcceptClient()
        If client = winsock2.INVALID_SOCKET Then Exit Do
        Set connection = New HttpConnection
        connection.Init client
        state.connections.Add connection
    Loop
End Sub

Private Sub ServiceConnections()
    Dim i As Long, connection As HttpConnection
    For i = state.connections.Count To 1 Step -1
        Set connection = state.connections(i)
        If connection.Service(state.appRouter) Then state.connections.Remove i
    Next
End Sub

Private Sub ScheduleTick()
    state.nextTick = Now + TimeSerial(0, 0, 1)
    Application.OnTime state.nextTick, TickProcedure()
    state.tickScheduled = True
End Sub

Private Sub CancelTick()
    If Not state.tickScheduled Then Exit Sub
    On Error Resume Next 'already fired or never registered
    Application.OnTime state.nextTick, TickProcedure(), , False
    On Error GoTo 0
    state.tickScheduled = False
End Sub

'Workbook-qualified so the right macro runs even with other workbooks open.
Private Function TickProcedure() As String
    TickProcedure = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & TICK_PROC
End Function

Private Sub SetStrayTickMarker(ByVal value As String)
    SetEnvironmentVariableW StrPtr(STRAY_TICK_VARIABLE), StrPtr(value)
End Sub

Private Sub ClearStrayTickMarker()
    SetEnvironmentVariableW StrPtr(STRAY_TICK_VARIABLE), 0
End Sub

Private Function GetStrayTickMarker() As String
    Dim size As Long
    size = GetEnvironmentVariableW(StrPtr(STRAY_TICK_VARIABLE), 0, 0)
    If size <= 1 Then Exit Function
    Dim buffer As String
    buffer = String$(size, vbNullChar)
    size = GetEnvironmentVariableW(StrPtr(STRAY_TICK_VARIABLE), StrPtr(buffer), size)
    GetStrayTickMarker = Left$(buffer, size)
End Function

Private Function SecondsSince(ByVal startedAt As Double) As Double
    SecondsSince = Timer - startedAt
    If SecondsSince < 0 Then SecondsSince = SecondsSince + 86400
End Function
