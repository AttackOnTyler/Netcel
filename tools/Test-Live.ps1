<#
.SYNOPSIS
  Builds a workbook from src/ via COM, starts the Netcel HTTP server inside Excel,
  exercises it with curl, then shuts everything down.

.DESCRIPTION
  Requires Excel (desktop) and "Trust access to the VBA project object model"
  (HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1). This script
  never changes that setting; it exits with code 2 if it is off.

  Only the Excel process this script starts is ever killed. A watchdog job kills it
  after -TimeoutSeconds in case a COM call hangs (e.g. a modal VBA error dialog).

  Output workbook: build/Netcel.xlsm (git-ignored).
#>
param(
    [int]$Port = 8765,
    [int]$TimeoutSeconds = 120
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$src = Join-Path $repo 'src'
$buildDir = Join-Path $repo 'build'
$workbookPath = Join-Path $buildDir 'Netcel.xlsm'

$security = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -ErrorAction SilentlyContinue
if (-not $security -or $security.AccessVBOM -ne 1) {
    Write-Host 'AccessVBOM is not enabled. In Excel: File > Options > Trust Center > Trust Center Settings >'
    Write-Host 'Macro Settings > "Trust access to the VBA project object model". Not changing it automatically.'
    exit 2
}

if (Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue) {
    throw "Port $Port is already in use; pass -Port."
}

Add-Type -Namespace NetcelTest -Name User32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
'@

New-Item -ItemType Directory -Force $buildDir | Out-Null

$script:failures = 0
function Assert-That([bool]$condition, [string]$message) {
    if ($condition) { Write-Host "PASS  $message" }
    else { Write-Host "FAIL  $message"; $script:failures++ }
}

function Invoke-Http([string[]]$curlArgs) {
    $marker = '__NETCEL_STATUS__'
    $output = & curl.exe -s --max-time 5 -w "$marker%{http_code}" @curlArgs
    $text = ($output -join "`n")
    $index = $text.LastIndexOf($marker)
    if ($index -lt 0) { return [pscustomobject]@{ Status = 0; Body = $text; ExitCode = $LASTEXITCODE } }
    [pscustomobject]@{
        Status   = [int]$text.Substring($index + $marker.Length)
        Body     = $text.Substring(0, $index)
        ExitCode = $LASTEXITCODE
    }
}

function Test-Listening { [bool](Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue) }

$xl = $null
$workbook = $null
[uint32]$excelPid = 0
$watchdog = $null

try {
    $xl = New-Object -ComObject Excel.Application
    [void][NetcelTest.User32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$excelPid)
    Write-Host "Started Excel PID $excelPid ($($xl.Version))"
    $watchdog = Start-Job -ScriptBlock {
        param($processId, $seconds)
        Start-Sleep -Seconds $seconds
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
    } -ArgumentList $excelPid, $TimeoutSeconds

    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    $workbook = $xl.Workbooks.Add()
    $project = $workbook.VBProject

    Get-ChildItem "$src\*" -Include *.bas, *.cls |
        Where-Object { $_.Name -notlike 'Test*' } |   # Rubberduck test modules need the Rubberduck reference
        ForEach-Object { [void]$project.VBComponents.Import($_.FullName) }

    $thisWorkbookCode = (Get-Content (Join-Path $src 'ThisWorkbook.doccls') -Raw -Encoding UTF8).TrimStart([char]0xFEFF)
    $codeModule = $project.VBComponents.Item($workbook.CodeName).CodeModule
    if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }
    $codeModule.AddFromString($thisWorkbookCode)

    if (Test-Path $workbookPath) { Remove-Item $workbookPath -Force }
    $workbook.SaveAs($workbookPath, 52) # xlOpenXMLWorkbookMacroEnabled
    $macroPrefix = "'$($workbook.Name)'!"
    Write-Host "Built $workbookPath"

    $started = $xl.Run("${macroPrefix}NetcelServer.StartServer", $Port)
    Assert-That ($started -eq $true) "StartServer($Port) returned True"
    Assert-That (Test-Listening) "socket is listening on $Port"

    $ping = Invoke-Http @("http://127.0.0.1:$Port/ping")
    Assert-That ($ping.Status -eq 200 -and $ping.Body -eq 'pong') "GET /ping -> $($ping.Status) '$($ping.Body)'"

    $headers = & curl.exe -s -i --max-time 5 "http://127.0.0.1:$Port/ping"
    Write-Host ($headers -join "`n")

    $eval = Invoke-Http @('-X', 'POST', '--data', '1+1', "http://127.0.0.1:$Port/eval")
    Assert-That ($eval.Status -eq 501) "POST /eval -> $($eval.Status) '$($eval.Body.Trim())'"

    $missing = Invoke-Http @("http://127.0.0.1:$Port/nope")
    Assert-That ($missing.Status -eq 404) "GET /nope -> $($missing.Status)"

    $wrongMethod = Invoke-Http @('-X', 'POST', "http://127.0.0.1:$Port/ping")
    Assert-That ($wrongMethod.Status -eq 405) "POST /ping -> $($wrongMethod.Status)"

    # Excel must stay responsive while serving: COM calls into it should return quickly.
    $timer = [Diagnostics.Stopwatch]::StartNew()
    for ($i = 1; $i -le 5; $i++) {
        $workbook.Worksheets.Item(1).Range('A1').Value2 = $i
        Start-Sleep -Milliseconds 300
    }
    $timer.Stop()
    $overhead = $timer.ElapsedMilliseconds - 1500
    Assert-That ($overhead -lt 1500) "Excel responsive while server running (5 COM writes, ${overhead}ms overhead)"

    $burst = 1..5 | ForEach-Object { Invoke-Http @("http://127.0.0.1:$Port/ping") }
    Assert-That (@($burst | Where-Object { $_.Status -eq 200 -and $_.Body -eq 'pong' }).Count -eq 5) "5 sequential /ping requests all succeed"

    [void]$xl.Run("${macroPrefix}NetcelServer.StopServer")
    Assert-That (-not (Test-Listening)) "StopServer closed the listening socket"
    $afterStop = Invoke-Http @("http://127.0.0.1:$Port/ping")
    Assert-That ($afterStop.Status -eq 0) "requests fail after StopServer (curl exit $($afterStop.ExitCode))"

    # Clean shutdown on workbook close: restart, then close the workbook without calling StopServer.
    $restarted = $xl.Run("${macroPrefix}NetcelServer.StartServer", $Port)
    Assert-That ($restarted -eq $true -and (Test-Listening)) "server restarts on the same port"
    $workbook.Close($false)
    $workbook = $null
    Start-Sleep -Seconds 4   # a stray OnTime would reopen the workbook; the guard should close it again
    Assert-That (-not (Test-Listening)) "closing the workbook closed the listening socket"
    $openNames = @(foreach ($book in $xl.Workbooks) { $book.Name })
    Assert-That ($openNames -notcontains 'Netcel.xlsm') "workbook not left open by a stray OnTime (open: $($openNames -join ', '))"
}
catch {
    Write-Host "ERROR $($_.Exception.Message)"
    $script:failures++
}
finally {
    if ($workbook) { try { $workbook.Close($false) } catch { } }
    if ($xl) { try { $xl.Quit() } catch { } }
    # Excel only exits once every COM reference held by this script is released.
    foreach ($name in 'book', 'codeModule', 'project', 'workbook', 'xl') {
        $value = Get-Variable -Name $name -ValueOnly -ErrorAction SilentlyContinue
        if ($value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch { }
        }
        Remove-Variable -Name $name -ErrorAction SilentlyContinue
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()

    if ($excelPid -ne 0) {
        $process = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($process -and -not $process.WaitForExit(10000)) {
            Write-Host "Excel PID $excelPid did not exit; killing it"
            Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
        }
    }
    if ($watchdog) { Stop-Job $watchdog -ErrorAction SilentlyContinue; Remove-Job $watchdog -Force -ErrorAction SilentlyContinue }
}

if ($script:failures -gt 0) { Write-Host "$script:failures check(s) failed"; exit 1 }
Write-Host 'All checks passed'
exit 0
