param([Parameter(Mandatory=$true)][string]$SessionDir)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$expectedRoot=Join-Path $env:TEMP 'AirDuctLaunch_'
$resolvedSession=[IO.Path]::GetFullPath($SessionDir)
if (-not $resolvedSession.StartsWith($expectedRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'Not a launcher test session' }
$runtimeMacro=Join-Path $resolvedSession 'generator.swp'
$contextFile=Join-Path $resolvedSession 'context.ini'
if (-not (Test-Path -LiteralPath $contextFile)) { throw 'Test context absent' }
$contextText=Get-Content -LiteralPath $contextFile -Raw
$expectedOutput = Get-ExternalTestPath 'paper_formats\A4.landscape\out'
$outputLine = @($contextText -split '\r?\n' | Where-Object { $_.StartsWith('Output=', [StringComparison]::OrdinalIgnoreCase) })
if ($outputLine.Count -ne 1 -or -not $outputLine[0].Substring(7).Equals($expectedOutput, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Not the dedicated test output'
}
$ownedProcesses=@(Get-CimInstance Win32_Process -Filter "Name='SLDWORKS.exe'" | Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($runtimeMacro,[StringComparison]::OrdinalIgnoreCase) -ge 0 })
if ($ownedProcesses.Count -ne 1) { throw 'Exactly one owned native process was not confirmed' }
$ownedProcess=$ownedProcesses[0]
$result=Invoke-CimMethod -InputObject $ownedProcess -MethodName Terminate -Arguments @{Reason=[uint32]85}
if ($result.ReturnValue -ne 0) { throw "Termination failed: $($result.ReturnValue)" }
@{pid=$ownedProcess.ProcessId; reason=85; session=$resolvedSession} | ConvertTo-Json -Compress
