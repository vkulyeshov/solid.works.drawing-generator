# Shared external paths for native tests. No generated data belongs in Git.
$TestRepository = Split-Path $PSScriptRoot
function Get-ExternalTestPath([string]$RelativePath) {
    $testBase = $env:SW_DRAWING_WORK_DIR
    if ([string]::IsNullOrWhiteSpace($testBase)) {
        $testBase = Join-Path (Split-Path $TestRepository) 'tmp\solid.works.drawing-generator'
    }
    if ([IO.Path]::IsPathRooted($RelativePath)) { $testPath = [IO.Path]::GetFullPath($RelativePath) }
    else { $testPath = [IO.Path]::GetFullPath((Join-Path $testBase $RelativePath)) }
    $testRoot = [IO.Path]::GetFullPath($TestRepository)
    if ($testPath -eq $testRoot -or $testPath.StartsWith($testRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Generated test data must stay outside the repository'
    }
    return $testPath
}
