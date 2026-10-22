#requires -Version 7.0
# Requires an up-to-date restore. Reads only projects listed by the solution.
# NuGet restore performs severity-based auditing; OSV independently checks every
# resolved package, including when NuGet publishes stale vulnerability metadata.
[CmdletBinding()]
param(
    [string]$SolutionPath = (Join-Path $PSScriptRoot '../TemplateFrame.slnx'),
    [string]$ReportPath = (Join-Path $PSScriptRoot '../artifacts/dependency-audit.json'),
    [ValidateRange(1, 365)]
    [int]$MaxFeedAgeDays = 7
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$solutionFile = (Resolve-Path -LiteralPath $SolutionPath).Path
$solutionRoot = Split-Path -Parent $solutionFile
[xml]$solution = Get-Content -LiteralPath $solutionFile -Raw
$projects = @($solution.SelectNodes('//Project[@Path]'))
if ($projects.Count -eq 0) {
    throw "No projects found in $solutionFile. Expected a .slnx solution."
}

$packageMap = @{}
foreach ($project in $projects) {
    $projectFile = [IO.Path]::GetFullPath((Join-Path $solutionRoot $project.GetAttribute('Path')))
    if (-not (Test-Path -LiteralPath $projectFile -PathType Leaf)) {
        throw "Solution project does not exist: $projectFile"
    }
    $assetsFile = Join-Path (Split-Path -Parent $projectFile) 'obj/project.assets.json'
    if (-not (Test-Path -LiteralPath $assetsFile -PathType Leaf)) {
        throw "Missing restore assets: $assetsFile. Run dotnet restore TemplateFrame.slnx first."
    }
    $assets = Get-Content -LiteralPath $assetsFile -Raw | ConvertFrom-Json -AsHashtable
    if (-not $assets.ContainsKey('libraries') -or -not $assets.ContainsKey('project') -or
        -not $assets.project.ContainsKey('restore') -or
        -not $assets.project.restore.ContainsKey('projectPath')) {
        throw "Incomplete restore assets: $assetsFile"
    }
    $restoredProject = [IO.Path]::GetFullPath($assets.project.restore.projectPath)
    $pathComparison = if ($IsWindows) { [StringComparison]::OrdinalIgnoreCase } else { [StringComparison]::Ordinal }
    if (-not $projectFile.Equals($restoredProject, $pathComparison)) {
        throw "Assets belong to a different project: $assetsFile"
    }
    foreach ($library in $assets.libraries.GetEnumerator()) {
        if ($library.Value.type -ne 'package') { continue }
        $separator = $library.Key.LastIndexOf('/')
        if ($separator -le 0 -or $separator -eq $library.Key.Length - 1) {
            throw "Invalid package entry '$($library.Key)' in $assetsFile"
        }
        $id = $library.Key.Substring(0, $separator)
        $version = $library.Key.Substring($separator + 1)
        $packageMap["$id/$version"] = [ordered]@{ id = $id; version = $version }
    }
}
$packages = @($packageMap.Values | Sort-Object { $_.id }, { $_.version })
if ($packages.Count -eq 0) { throw 'No resolved packages found; refusing an empty dependency audit.' }
Write-Host "Auditing $($packages.Count) resolved package versions from $($projects.Count) solution projects."

$nugetIndexUrl = 'https://api.nuget.org/v3/vulnerabilities/index.json'
$nugetIndex = Invoke-RestMethod -Uri $nugetIndexUrl -TimeoutSec 30 -MaximumRetryCount 2 -RetryIntervalSec 1
if ($nugetIndex.Count -eq 0) { throw 'NuGet vulnerability index is empty.' }
$updatedDates = foreach ($entry in $nugetIndex) {
    $updatedProperty = $entry.PSObject.Properties['@updated']
    if ($null -eq $updatedProperty -or [string]::IsNullOrWhiteSpace($updatedProperty.Value)) {
        throw 'NuGet vulnerability index has no usable update timestamp.'
    }
    if ($updatedProperty.Value -is [DateTime]) {
        [DateTimeOffset]$updatedProperty.Value
    }
    else {
        [DateTimeOffset]::Parse($updatedProperty.Value, [Globalization.CultureInfo]::InvariantCulture)
    }
}
# The update partition extends the base partition; use the newest valid timestamp.
$feedUpdated = $updatedDates | Sort-Object -Descending | Select-Object -First 1
$now = [DateTimeOffset]::UtcNow
if ($feedUpdated -gt $now.AddDays(1)) { throw 'NuGet vulnerability index has a future update timestamp.' }
$feedAgeDays = ($now - $feedUpdated).TotalDays
$feedStale = $feedAgeDays -gt $MaxFeedAgeDays
if ($feedStale) {
    Write-Warning "NuGet vulnerability metadata is stale (newest update $($feedUpdated.ToString('u'))). A clean restore is insufficient; the required independent OSV audit must succeed."
}

$osvUrl = 'https://api.osv.dev/v1/querybatch'
# A known vulnerable version is a positive control for a usable OSV response.
$control = [ordered]@{ id = 'System.IO.Packaging'; version = '8.0.0' }
$queries = @($packages) + @($control)
$findings = [Collections.Generic.List[object]]::new()
$controlPassed = $false
$batchSize = 1000
for ($offset = 0; $offset -lt $queries.Count; $offset += $batchSize) {
    $end = [Math]::Min($offset + $batchSize, $queries.Count)
    $batch = @($queries[$offset..($end - 1)])
    $body = @{ queries = @($batch | ForEach-Object {
        @{ package = @{ ecosystem = 'NuGet'; name = $_.id }; version = $_.version }
    }) } | ConvertTo-Json -Depth 6 -Compress
    $response = Invoke-RestMethod -Method Post -Uri $osvUrl -Body $body -ContentType 'application/json' -TimeoutSec 30 -MaximumRetryCount 2 -RetryIntervalSec 1
    $resultProperty = $response.PSObject.Properties['results']
    if ($null -eq $resultProperty -or @($resultProperty.Value).Count -ne $batch.Count) {
        throw "OSV returned an incomplete batch at offset $offset."
    }
    $results = @($resultProperty.Value)
    for ($index = 0; $index -lt $batch.Count; $index++) {
        $result = $results[$index]
        if ($null -eq $result) { throw 'OSV returned a null query result.' }
        if ($null -ne $result.PSObject.Properties['error']) { throw 'OSV returned a per-package query error.' }
        $vulnsProperty = $result.PSObject.Properties['vulns']
        $vulnerabilities = if ($null -eq $vulnsProperty) { @() } else { @($vulnsProperty.Value) }
        $ids = @($vulnerabilities | ForEach-Object {
            $idProperty = $_.PSObject.Properties['id']
            if ($null -eq $idProperty -or [string]::IsNullOrWhiteSpace($idProperty.Value)) {
                throw 'OSV returned a vulnerability without an identifier.'
            }
            $idProperty.Value
        })
        $pageProperty = $result.PSObject.Properties['next_page_token']
        if ($null -ne $pageProperty -and $pageProperty.Value -and $ids.Count -eq 0) {
            throw 'OSV returned pagination without any vulnerability identifiers.'
        }
        if ($offset + $index -eq $packages.Count) {
            $controlPassed = $ids -contains 'GHSA-f32c-w444-8ppv' -or $ids -contains 'GHSA-qj66-m88j-hmgj'
            if (-not $controlPassed) { throw 'OSV did not detect the known vulnerable positive control.' }
        }
        elseif ($ids.Count -gt 0) {
            $findings.Add([ordered]@{
                package = $batch[$index].id
                version = $batch[$index].version
                vulnerability_ids = $ids
            })
        }
    }
}
if (-not $controlPassed) { throw 'OSV positive control was not checked.' }

$report = [ordered]@{
    checked_at_utc = $now.ToString('o')
    solution = [IO.Path]::GetFileName($solutionFile)
    project_count = $projects.Count
    package_count = $packages.Count
    packages = $packages
    nuget = @{ source = $nugetIndexUrl; newest_update_utc = $feedUpdated.ToString('o'); stale = $feedStale; max_age_days = $MaxFeedAgeDays }
    osv = @{ source = $osvUrl; positive_control_passed = $controlPassed; findings = @($findings.ToArray()) }
}
$reportFile = [IO.Path]::GetFullPath($ReportPath)
[IO.Directory]::CreateDirectory((Split-Path -Parent $reportFile)) | Out-Null
$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $reportFile -Encoding utf8
Write-Host "Dependency audit report: $reportFile"
if ($findings.Count -gt 0) {
    foreach ($finding in $findings) {
        Write-Host "$($finding.package) $($finding.version): $($finding.vulnerability_ids -join ', ')"
    }
    throw "OSV found vulnerable versions in $($findings.Count) resolved package entries."
}
Write-Host "OSV audit passed for $($packages.Count) resolved package versions; positive control passed."
