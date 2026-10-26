param(
    [string]$ResultsRoot = "$PSScriptRoot/raw",
    [string]$OutputPath = "$PSScriptRoot/summary.csv"
)

$ErrorActionPreference = 'Stop'
$rows = foreach ($file in Get-ChildItem -LiteralPath $ResultsRoot -Filter '*full-compressed.json' -Recurse | Sort-Object FullName) {
    $report = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
    foreach ($benchmark in $report.Benchmarks) {
        if ($null -eq $benchmark.Statistics) { throw "Missing statistics: $($file.FullName)" }
        [pscustomobject]@{
            Run = $file.Directory.Parent.Name
            Method = $benchmark.Method
            Parameters = $benchmark.Parameters
            N = $benchmark.Statistics.N
            MeanMs = $benchmark.Statistics.Mean / 1e6
            ErrorMs = $benchmark.Statistics.ConfidenceInterval.Margin / 1e6
            StdDevMs = $benchmark.Statistics.StandardDeviation / 1e6
            AllocatedBytes = $benchmark.Memory.BytesAllocatedPerOperation
        }
    }
}
$rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
$rows
