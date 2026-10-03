Attribute VB_Name = "app"
'@Folder("app")
Option Explicit

'Starts the HTTP server on the default port (127.0.0.1:8080).
Public Sub main()
    NetcelServer.StartServer
End Sub

Public Sub StopMain()
    NetcelServer.StopServer
End Sub
