param(
    [Parameter(Mandatory)][string]$ResultDirectory,
    [Parameter(Mandatory)][string]$OutputCsv,
    [string]$ExcludedCasesPath
)
$ErrorActionPreference = 'Stop'
$excluded = @{}
if ($ExcludedCasesPath) {
    foreach ($item in Import-Csv -LiteralPath $ExcludedCasesPath) { $excluded[$item.Case] = $item.Reason }
}
$rows = foreach ($file in Get-ChildItem -LiteralPath $ResultDirectory -Filter summary.json -Recurse) {
    $summary = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
    $events = @(Get-Content -LiteralPath (Join-Path $file.DirectoryName 'stdout.jsonl') |
        Where-Object { $_.StartsWith('{') } | ForEach-Object { $_ | ConvertFrom-Json })
    $start = $events | Where-Object Phase -eq 'prepare' | Select-Object -First 1
    $operation = $events | Where-Object Phase -eq 'validation' | Select-Object -Last 1
    $finish = $events | Where-Object Phase -eq 'complete' | Select-Object -Last 1
    $status = if ($excluded.ContainsKey($file.Directory.Name)) { $excluded[$file.Directory.Name] }
    elseif ($summary.Guard) { $summary.Guard }
    elseif ($finish.Error -like 'System.OutOfMemoryException:*') { 'out-of-memory' }
    elseif ($finish.Status) { $finish.Status } else { 'process-failed' }
    [pscustomobject]@{
        Case = $file.Directory.Name
        Mode = $summary.Mode
        Rows = $summary.Rows
        Columns = $summary.Columns
        Concurrency = $summary.Concurrency
        Framework = $summary.Framework
        Runtime = $start.Runtime
        Limits = $summary.Limits
        Status = $status
        WallSeconds = $summary.WallSeconds
        BatchSeconds = if ($operation) { $operation.BatchMilliseconds / 1000 } else { $null }
        SlowestSeconds = if ($operation) { $operation.MaximumMilliseconds / 1000 } else { $null }
        SampledPeakPrivateMiB = $summary.SampledPeakPrivateBytes / 1MB
        SampledPeakWorkingMiB = $summary.SampledPeakWorkingSetBytes / 1MB
        OsPeakWorkingMiB = if ($finish.PeakWorkingSetBytes) { $finish.PeakWorkingSetBytes / 1MB } else { $null }
        MinimumFreeMiB = $summary.MinimumFreeBytes / 1MB
        OutputMiB = if ($finish) { $finish.OutputBytes / 1MB } else { $null }
        Error = if ($finish.Error) { ($finish.Error -split "`r?`n")[0] } else { $null }
    }
}
$rows | Sort-Object Framework, Mode, Rows, Columns, Concurrency, Case | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation
$rows | Group-Object Status | Select-Object Name, Count
