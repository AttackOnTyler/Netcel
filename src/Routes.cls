VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Routes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("app.http")
Option Explicit

'Route handlers. Each must be Public and match:
'    Function X(ByVal request As HttpRequest) As HttpResponse

Public Sub Register(ByVal target As Router)
    target.AddRoute "GET", "/ping", Me, "Ping"
    target.AddRoute "POST", "/eval", Me, "Eval"
End Sub

Public Function Ping(ByVal request As HttpRequest) As HttpResponse
    Set Ping = Http.NewResponse(200, "pong")
End Function

'Placeholder for evaluating VBA/VBish commands against the live workbook.
'Deliberately unimplemented: it will need auth before it is safe to enable.
Public Function Eval(ByVal request As HttpRequest) As HttpResponse
    Set Eval = Http.NewResponse(501, "eval is not implemented yet" & vbLf)
End Function
