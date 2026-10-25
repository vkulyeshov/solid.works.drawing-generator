param([switch]$VerifyOnly, [string]$Only = '', [string]$OutputDirectory = '')
$ErrorActionPreference = 'Stop'
$repository = Split-Path (Split-Path $PSScriptRoot)
if ($OutputDirectory -eq '') {
    $workBase = $env:SW_DRAWING_WORK_DIR
    if ([string]::IsNullOrWhiteSpace($workBase)) {
        $workBase = Join-Path (Split-Path $repository) 'tmp\solid.works.drawing-generator'
    }
    $OutputDirectory = Join-Path $workBase 'templates'
}
$output = [IO.Path]::GetFullPath($OutputDirectory)
if ($output -eq $repository -or $output.StartsWith($repository + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Template development outputs must stay outside the repository'
}
$api = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist'
Add-Type -Path "$api\SolidWorks.Interop.sldworks.dll"
Add-Type -Path "$api\SolidWorks.Interop.swconst.dll"
Add-Type -TypeDefinition (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'CreateDstuTemplates.cs') -Raw -Encoding UTF8) -ReferencedAssemblies "$api\SolidWorks.Interop.sldworks.dll","$api\SolidWorks.Interop.swconst.dll",'System.Web.Extensions.dll'
$reportPath = Join-Path $output 'Verification\native-report.json'
$oldRecords = $null
if ($Only -ne '') {
    if ($Only -notin @('DSTU_A4_Portrait','DSTU_A4_Landscape','DSTU_A3_Portrait','DSTU_A3_Landscape')) { throw 'Unknown template name' }
    if (Test-Path -LiteralPath $reportPath) { $oldRecords = [System.IO.File]::ReadAllText($reportPath) }
}
[DstuTemplateBuilder]::Run($output, $VerifyOnly.IsPresent, $Only)
if ($null -ne $oldRecords) {
    $jsonReader = [System.Web.Script.Serialization.JavaScriptSerializer]::new()
    $jsonReader.MaxJsonLength = 4000000
    $newRecords = $jsonReader.DeserializeObject([System.IO.File]::ReadAllText($reportPath))
    $merged = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $jsonReader.DeserializeObject($oldRecords)) {
        if ($record['name'] -eq $Only) { foreach ($replacement in $newRecords) { $merged.Add($replacement) } }
        else { $merged.Add($record) }
    }
    [System.IO.File]::WriteAllText($reportPath, $jsonReader.Serialize($merged), [System.Text.UTF8Encoding]::new($false))
}
