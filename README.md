# Netcel

An HTTP server that runs inside Excel, written in VBA on top of Winsock2.

The goal is an HTTP (and later MCP) endpoint in a live workbook, so that tools and AI
agents can call into it. [VBish](https://github.com/AttackOnTyler/VBish) commands are
meant to become routes.

## Status

- `GET /ping` returns `200 pong`
- `POST /eval` returns `501 Not Implemented` (a stub; it needs auth before it can be enabled)
- Unknown paths return `404`; a known path with the wrong method returns `405`
- Binds to `127.0.0.1` by default

## How it works

- Sockets are non-blocking (`ioctlsocket FIONBIO`), so nothing ever waits inside Winsock.
- `Application.OnTime` calls `NetcelServer.ServerTick` about once a second. Each tick
  accepts pending connections and services them for at most 250 ms, then returns
  control to Excel.
- Each connection handles one request and then closes (`Connection: close`).
- The `Declare PtrSafe` signatures are 64-bit correct: `SOCKET` is `LongPtr`, and
  buffers are `Byte` arrays.
- `Workbook_BeforeClose` stops the server: it cancels the tick, closes the sockets and
  calls `WSACleanup`.

## Usage

Import `src/*` into a macro-enabled workbook. The `Test*` modules are Rubberduck tests
and need the Rubberduck reference. Paste `ThisWorkbook.doccls` into `ThisWorkbook`. Then:

```vb
NetcelServer.StartServer 8080   ' or run the macro app.main
' curl http://127.0.0.1:8080/ping  ->  pong
NetcelServer.StopServer
```

Routes are registered in `Routes.Register`. A handler is a public method with this shape:
`Function X(ByVal request As HttpRequest) As HttpResponse`.

## Live test

`tools/Test-Live.ps1` builds `build/Netcel.xlsm` from `src/` through Excel COM, starts
the server, checks it with `curl`, then shuts everything down. It needs "Trust access to
the VBA project object model" to be enabled, and never enables it for you.

```powershell
powershell -ExecutionPolicy Bypass -File tools\Test-Live.ps1 [-Port 8765]
```
