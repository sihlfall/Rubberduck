[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][int]$HostProcessId,
    [ValidateSet('describe', 'run', 'cancel')][string]$Command = 'describe',
    [string]$RequestFile,
    [string]$RunId,
    [int]$ConnectTimeoutMs = 5000
)
$ErrorActionPreference = 'Stop'
if ($Command -eq 'run') {
    if (-not $RequestFile) { throw 'run requires -RequestFile.' }
    $requestText = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $RequestFile).Path)
    $request = $requestText | ConvertFrom-Json
    if ($request.version -ne 1 -or $request.command -ne 'run') { throw 'Expected a version 1 run request.' }
} else {
    if ($Command -eq 'cancel' -and -not $RunId) { throw 'cancel requires -RunId.' }
    $requestText = @{version=1; command=$Command; runId=$RunId} | ConvertTo-Json -Compress
}
$payload = [Text.Encoding]::UTF8.GetBytes($requestText)
if ($payload.Length -gt 1048576) { throw 'Request exceeds the 1 MiB protocol limit.' }
$pipe = New-Object IO.Pipes.NamedPipeClientStream('.', "Rubberduck.Tests.$HostProcessId", [IO.Pipes.PipeDirection]::InOut, [IO.Pipes.PipeOptions]::Asynchronous)
try {
    $pipe.Connect($ConnectTimeoutMs)
    $writer = New-Object IO.BinaryWriter($pipe)
    $reader = New-Object IO.BinaryReader($pipe)
    $writer.Write([int]$payload.Length)
    $writer.Write($payload)
    $writer.Flush()
    $failed = $false
    while ($true) {
        $length = $reader.ReadInt32()
        if ($length -lt 1 -or $length -gt 1048576) { throw 'Invalid response frame length.' }
        $bytes = $reader.ReadBytes($length)
        if ($bytes.Length -ne $length) { throw 'Incomplete response frame.' }
        $text = [Text.Encoding]::UTF8.GetString($bytes)
        Write-Output $text
        $testEvent = $text | ConvertFrom-Json
        if ($testEvent.event -eq 'error') { throw $testEvent.message }
        if ($testEvent.event -eq 'issue' -or ($testEvent.event -eq 'test_completed' -and $testEvent.status -notin @('Succeeded','Ignored'))) { $failed = $true }
        if ($testEvent.event -in @('description','cancel_requested')) { return }
        if ($testEvent.event -eq 'run_completed') {
            if ($testEvent.cancelled) { throw 'Test run was cancelled.' }
            if ($failed) { throw 'Test run completed with failures or execution issues.' }
            return
        }
    }
} finally {
    $pipe.Dispose()
}
