param([Parameter(Mandatory=$true)][string]$WorkDirectory,
      [string]$CaseFilter = '',
      [string[]]$Cases = @('matochina02','rightside01','roof02','sheetmetal01','a3landscape01'))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$work = Get-ExternalTestPath $WorkDirectory
$results = @()
if ($CaseFilter) {
    $Cases = $CaseFilter.Split(',')
}
foreach ($case in $Cases) {
    $directory = Join-Path $work $case
    $settings = Get-Content -LiteralPath (Join-Path $directory 'generator.ini') -Encoding Unicode
    $inputSetting = @($settings | Where-Object { $_ -match '^LauncherInput=' })
    if ($inputSetting.Count -ne 1) {
        throw "The prepared fixture needs one LauncherInput setting: $case"
    }
    $modelName = [IO.Path]::GetFileName($inputSetting[0].Substring('LauncherInput='.Length))
    if (-not (Test-Path -LiteralPath (Join-Path (Join-Path $directory 'input') $modelName))) {
        throw "The prepared fixture input is missing: $case"
    }
    Write-Output "Memory layout case: $case"
    $nativeArguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot 'RunSixViewGeneration.ps1'),'-WorkDirectory',$directory,'-ModelFileName',$modelName,'-MacroTimeoutSeconds','240')
    if ([IO.Path]::GetExtension($modelName) -ieq '.SLDASM') {
        $nativeArguments += '-RemapInputReferences'
    }
    & "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" @nativeArguments
    $results += [pscustomobject]@{ Case=$case; ExitCode=$LASTEXITCODE }
}
$results | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $work 'native-suite.json') -Encoding UTF8
$results | Format-Table -AutoSize
if (@($results | Where-Object ExitCode -ne 0).Count -gt 0) {
    exit 1
}
