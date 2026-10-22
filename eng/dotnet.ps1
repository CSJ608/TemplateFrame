# Use the project-local patched host when present, including for test child processes.
[string[]]$DotnetArguments = $args

$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$localRoot = Join-Path $projectRoot '.dotnet'
$binaryName = if ($IsWindows) { 'dotnet.exe' } else { 'dotnet' }
$localBinary = Join-Path $localRoot $binaryName
if (-not (Test-Path -LiteralPath $localBinary -PathType Leaf)) {
    & dotnet @DotnetArguments
    exit $LASTEXITCODE
}

$previousRoot = $env:DOTNET_ROOT
$previousX64Root = $env:DOTNET_ROOT_X64
$previousPath = $env:PATH
$exitCode = 1
try {
    $env:DOTNET_ROOT = $localRoot
    $env:DOTNET_ROOT_X64 = $localRoot
    $env:PATH = $localRoot + [IO.Path]::PathSeparator + $previousPath
    & $localBinary @DotnetArguments
    $exitCode = $LASTEXITCODE
}
finally {
    $env:DOTNET_ROOT = $previousRoot
    $env:DOTNET_ROOT_X64 = $previousX64Root
    $env:PATH = $previousPath
}
exit $exitCode
