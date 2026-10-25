param(
    [Parameter(Mandatory)][string]$Mode,
    [Parameter(Mandatory)][int]$Rows,
    [int]$Columns = 6,
    [int]$Concurrency = 1,
    [ValidateSet('net8.0', 'net10.0')][string]$Framework = 'net10.0',
    [ValidateSet('default', 'capacity')][string]$Limits = 'capacity',
    [Parameter(Mandatory)][string]$CaseDirectory,
    [string]$FixtureDirectory,
    [int]$TimeoutSeconds = 180,
    [double]$MaxPrivateGiB = 6,
    [double]$MinimumFreeGiB = 4
)
$ErrorActionPreference = 'Stop'
$casePath = [IO.Path]::GetFullPath($CaseDirectory)
if (Test-Path -LiteralPath $casePath) { throw "Case directory already exists: $casePath" }
New-Item -ItemType Directory -Path $casePath | Out-Null
if ($Mode.EndsWith('read')) {
    $source = Join-Path $FixtureDirectory "input-$Rows-$Columns.xlsx"
    Copy-Item -LiteralPath $source -Destination $casePath
}
$assembly = Join-Path $PSScriptRoot "bin/Release/$Framework/TemplateFrame.CapacityProbe.dll"
$standardOutput = Join-Path $casePath 'stdout.jsonl'
$standardError = Join-Path $casePath 'stderr.txt'
$start = [Diagnostics.ProcessStartInfo]::new('dotnet')
$start.UseShellExecute = $false
$start.CreateNoWindow = $true
$start.RedirectStandardOutput = $true
$start.RedirectStandardError = $true
foreach ($argument in @($assembly, $Mode, "$Rows", "$Columns", "$Concurrency", $casePath, $Limits)) {
    $start.ArgumentList.Add($argument)
}
$start.Environment['DOTNET_GCHeapHardLimit'] = '0x140000000'
$start.Environment.Remove('DOTNET_TieredCompilation') | Out-Null
$start.Environment.Remove('COMPlus_TieredCompilation') | Out-Null
$available = (Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory * 1KB
if ($available -lt ($MinimumFreeGiB + 1) * 1GB) { throw 'Insufficient free memory to start safely.' }
$process = [Diagnostics.Process]::new()
$process.StartInfo = $start
$timer = [Diagnostics.Stopwatch]::StartNew()
$process.Start() | Out-Null
$stdoutTask = $process.StandardOutput.ReadToEndAsync()
$stderrTask = $process.StandardError.ReadToEndAsync()
$peakPrivate = 0L
$peakWorking = 0L
$minimumFree = $available
$guard = $null
$sampleCount = 0
$lastFreeCheck = 0.0
$samples = [Collections.Generic.List[object]]::new()
try {
    while (-not $process.HasExited) {
        $process.Refresh()
        $private = $process.PrivateMemorySize64
        $working = $process.WorkingSet64
        $peakPrivate = [Math]::Max($peakPrivate, $private)
        $peakWorking = [Math]::Max($peakWorking, $working)
        if ($timer.Elapsed.TotalSeconds - $lastFreeCheck -ge 1) {
            $available = (Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory * 1KB
            $minimumFree = [Math]::Min($minimumFree, $available)
            $lastFreeCheck = $timer.Elapsed.TotalSeconds
        }
        $samples.Add([pscustomobject]@{Seconds=$timer.Elapsed.TotalSeconds;PrivateBytes=$private;WorkingSetBytes=$working;FreeBytes=$available})
        $sampleCount++
        if ($private -gt $MaxPrivateGiB * 1GB) { $guard = 'private-memory-limit' }
        elseif ($available -lt $MinimumFreeGiB * 1GB) { $guard = 'system-free-memory-limit' }
        elseif ($timer.Elapsed.TotalSeconds -gt $TimeoutSeconds) { $guard = 'timeout' }
        if ($guard) { $process.Kill($true); break }
        Start-Sleep -Milliseconds 100
    }
    $process.WaitForExit()
    $stdoutTask.Result | Set-Content -LiteralPath $standardOutput -Encoding utf8
    $stderrTask.Result | Set-Content -LiteralPath $standardError -Encoding utf8
    $samples | Export-Csv -LiteralPath (Join-Path $casePath 'memory.csv') -NoTypeInformation
    $summary = [pscustomobject]@{
        Mode=$Mode;Rows=$Rows;Columns=$Columns;Concurrency=$Concurrency;Framework=$Framework;Limits=$Limits
        ExitCode=$process.ExitCode;Guard=$guard;WallSeconds=$timer.Elapsed.TotalSeconds
        SampledPeakPrivateBytes=$peakPrivate;SampledPeakWorkingSetBytes=$peakWorking;MinimumFreeBytes=$minimumFree
        MaxPrivateGiB=$MaxPrivateGiB;MinimumFreeGiB=$MinimumFreeGiB;TimeoutSeconds=$TimeoutSeconds
        ManagedHeapHardLimitGiB=5;MemoryScope='process lifetime including preparation and validation'
    }
    $summary | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $casePath 'summary.json') -Encoding utf8
    $summary | ConvertTo-Json -Compress
    Get-Content -LiteralPath $standardOutput
} finally {
    if (-not $process.HasExited) { $process.Kill($true); $process.WaitForExit() }
    $process.Dispose()
}
