VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Router"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("app.http")
Option Explicit

'Exact-match (method, path) router. A handler is any object plus the name of a public
'method on it with the signature:  Function X(ByVal request As HttpRequest) As HttpResponse
'Unknown path -> 404; known path with another method -> 405 with an Allow header.

Private Const ROUTE_METHOD As Long = 0
Private Const ROUTE_PATH As Long = 1
Private Const ROUTE_HANDLER As Long = 2
Private Const ROUTE_PROC As Long = 3

Private m_routes As Collection

Private Sub Class_Initialize()
    Set m_routes = New Collection
End Sub

Public Sub AddRoute(ByVal httpMethod As String, ByVal path As String, ByVal handler As Object, ByVal procName As String)
    m_routes.Add Array(UCase$(httpMethod), path, handler, procName)
End Sub

Public Function Dispatch(ByVal request As HttpRequest) As HttpResponse
    Dim route As Variant, allowed As String

    For Each route In m_routes
        If StrComp(route(ROUTE_PATH), request.Path, vbBinaryCompare) = 0 Then
            If StrComp(route(ROUTE_METHOD), request.Method, vbBinaryCompare) = 0 Then
                Set Dispatch = CallByName(route(ROUTE_HANDLER), route(ROUTE_PROC), VbMethod, request)
                Exit Function
            End If
            If Len(allowed) > 0 Then allowed = allowed & ", "
            allowed = allowed & route(ROUTE_METHOD)
        End If
    Next

    If Len(allowed) > 0 Then
        Set Dispatch = Http.NewResponse(405, "method not allowed" & vbLf)
        Dispatch.AddHeader "Allow", allowed
    Else
        Set Dispatch = Http.NewResponse(404, "not found" & vbLf)
    End If
End Function
