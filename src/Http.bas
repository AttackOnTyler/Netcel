Attribute VB_Name = "Http"
'@Folder("app.http")
Option Explicit

'Convenience factory for responses.
Public Function NewResponse(ByVal statusCode As Long, Optional ByVal body As String = vbNullString, Optional ByVal contentType As String = "text/plain; charset=utf-8") As HttpResponse
    Dim response As HttpResponse
    Set response = New HttpResponse
    response.StatusCode = statusCode
    response.Body = body
    response.ContentType = contentType
    Set NewResponse = response
End Function

Public Function ReasonPhrase(ByVal statusCode As Long) As String
    Select Case statusCode
        Case 200: ReasonPhrase = "OK"
        Case 201: ReasonPhrase = "Created"
        Case 204: ReasonPhrase = "No Content"
        Case 400: ReasonPhrase = "Bad Request"
        Case 404: ReasonPhrase = "Not Found"
        Case 405: ReasonPhrase = "Method Not Allowed"
        Case 408: ReasonPhrase = "Request Timeout"
        Case 413: ReasonPhrase = "Content Too Large"
        Case 431: ReasonPhrase = "Request Header Fields Too Large"
        Case 500: ReasonPhrase = "Internal Server Error"
        Case 501: ReasonPhrase = "Not Implemented"
        Case 503: ReasonPhrase = "Service Unavailable"
        Case Else: ReasonPhrase = "Unknown"
    End Select
End Function
