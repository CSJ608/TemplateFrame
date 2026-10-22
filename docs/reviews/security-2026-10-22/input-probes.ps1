# Read-only repository probes: all synthetic Office packages remain in memory.
# Run in a NEW pwsh process so different package/build versions cannot be mixed:
# pwsh -NoProfile -File .\input-probes.ps1 -Configuration Release
# AllocatedBytes is cumulative current-thread managed allocation, NOT peak memory,
# retained heap size, native allocation, process working set, or a load benchmark.
[CmdletBinding()]
param(
    [string]$RepositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..')),
    [ValidateSet('Debug', 'Release')][string]$Configuration = 'Release',
    [ValidateRange(2, 1000000)][int]$SparseRow = 100000,
    [ValidateRange(2, 100000)][int]$NamedRangeEndRow = 10000,
    [string]$ResultsPath = (Join-Path $PSScriptRoot 'input-probes-results.json')
)

$ErrorActionPreference = 'Stop'
$repository = (Resolve-Path -LiteralPath $RepositoryRoot).Path
$excelAssemblyDirectory = Join-Path $repository "test/TemplateFrame.Excel.Tests/bin/$Configuration/net8.0"
$simpleAssemblyDirectory = Join-Path $repository "test/TemplateFrame.Excel.Simple.Tests/bin/$Configuration/net8.0"
$assemblyPaths = @(
    (Join-Path $excelAssemblyDirectory 'System.IO.Packaging.dll'),
    (Join-Path $excelAssemblyDirectory 'DocumentFormat.OpenXml.Framework.dll'),
    (Join-Path $excelAssemblyDirectory 'DocumentFormat.OpenXml.dll'),
    (Join-Path $excelAssemblyDirectory 'TemplateFrame.dll'),
    (Join-Path $excelAssemblyDirectory 'TemplateFrame.Excel.dll'),
    (Join-Path $simpleAssemblyDirectory 'TemplateFrame.Excel.Simple.dll')
)
$assemblyManifest = @(
    foreach ($assemblyPath in $assemblyPaths) {
        if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
            throw "Missing compiled artifact: $assemblyPath. Build the matching tests first; this probe never builds or modifies the repository."
        }
        $loaded = [Reflection.Assembly]::LoadFrom($assemblyPath)
        [ordered]@{
            path = $assemblyPath
            # LoadFrom may return an already loaded assembly from a different path.
            # Keep assemblyVersion as the loaded-version field for prior result compatibility.
            assemblyVersion = $loaded.GetName().Version.ToString()
            loadedAssemblyVersion = $loaded.GetName().Version.ToString()
            loadedAssemblyLocation = $loaded.Location
            diskAssemblyVersion = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath).Version.ToString()
            sha256 = (Get-FileHash -LiteralPath $assemblyPath -Algorithm SHA256).Hash
            lastWriteTimeUtc = (Get-Item -LiteralPath $assemblyPath).LastWriteTimeUtc.ToString('o')
        }
    }
)

function New-ProbeWorkbook {
    param(
        [ValidateSet('Simple', 'NamedRange', 'Dtd', 'ActiveContent')][string]$Kind,
        [int]$LastRow = 2
    )
    $stream = [IO.MemoryStream]::new()
    $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $true)
    $workbookContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
    $extraContentType = ''
    $extraWorkbookRelation = ''
    $definedNames = ''
    if ($Kind -eq 'ActiveContent') {
        $workbookContentType = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
        $extraContentType = '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
        $extraWorkbookRelation = '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
    }
    if ($Kind -eq 'NamedRange') {
        $definedNames = '<definedNames><definedName name="TF_Items_Value">Sheet1!$A$2:$A$' + $LastRow + '</definedName></definedNames>'
        $worksheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="A2" t="inlineStr"><is><t>Data</t></is></c></row></sheetData></worksheet>'
    } else {
        $worksheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row><row r="' + $LastRow + '"><c r="A' + $LastRow + '" t="inlineStr"><is><t>Data</t></is></c></row></sheetData></worksheet>'
    }
    if ($Kind -eq 'Dtd') {
        # Harmless literal declaration only: no external entity, URL, file, or expansion bomb.
        $worksheet = '<!DOCTYPE worksheet [<!ENTITY x "safe-probe">]>' + $worksheet
    }
    if ($Kind -eq 'ActiveContent') {
        $worksheet = $worksheet.Replace('</worksheet>', '<hyperlinks><hyperlink ref="A1" r:id="rIdExternal"/></hyperlinks></worksheet>')
    }
    $parts = [ordered]@{
        '[Content_Types].xml' = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="' + $workbookContentType + '"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' + $extraContentType + '</Types>'
        '_rels/.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' + $definedNames + '</workbook>'
        'xl/_rels/workbook.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' + $extraWorkbookRelation + '</Relationships>'
        'xl/worksheets/sheet1.xml' = $worksheet
    }
    if ($Kind -eq 'ActiveContent') {
        # Not a valid VBA project and contains NO executable code. No URL is accessed.
        $parts['xl/vbaProject.bin'] = 'BENIGN_INERT_MARKER_NOT_EXECUTABLE'
        $parts['xl/worksheets/_rels/sheet1.xml.rels'] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/safe-inert-probe" TargetMode="External"/></Relationships>'
    }
    try {
        foreach ($entry in $parts.GetEnumerator()) {
            $writer = [IO.StreamWriter]::new($archive.CreateEntry($entry.Key).Open())
            try { $writer.Write($entry.Value) } finally { $writer.Dispose() }
        }
    } finally {
        $archive.Dispose()
    }
    $stream.Position = 0
    return ,$stream
}

# A small warmup isolates most one-time JIT/assembly initialization from the bounded probes.
$warmup = New-ProbeWorkbook -Kind Simple
try { $null = [TemplateFrame.Excel.Simple.SimpleExcel]::Read($warmup, $null) } finally { $warmup.Dispose() }
$stream = New-ProbeWorkbook -Kind Simple -LastRow $SparseRow
try {
    $watch = [Diagnostics.Stopwatch]::new()
    $before = [GC]::GetAllocatedBytesForCurrentThread()
    $watch.Start()
    $parsed = [TemplateFrame.Excel.Simple.SimpleExcel]::Read($stream, $null)
    $watch.Stop()
    $allocated = [GC]::GetAllocatedBytesForCurrentThread() - $before
    $simpleResult = [ordered]@{
        lastRowIndex = $SparseRow; physicalDataRows = 1; inputBytes = $stream.Length
        rowsReturned = $parsed.Rows.Count; elapsedMilliseconds = $watch.ElapsedMilliseconds
        cumulativeThreadAllocatedBytes = $allocated
    }
    $parsed = $null
} finally { $stream.Dispose() }

$column = [TemplateFrame.Contract.TextElement]::new()
$column.Key = 'Value'
$table = [TemplateFrame.Contract.TableElement]::new()
$table.Key = 'Items'
$table.Columns = [TemplateFrame.Contract.TextElement[]]@($column)
$contract = [TemplateFrame.Contract.TemplateContract]::new()
$contract.Elements = [TemplateFrame.Contract.TemplateElement[]]@($table)
$warmup = New-ProbeWorkbook -Kind NamedRange
try { $null = [TemplateFrame.Excel.ExcelTemplateParser]::new().Parse($warmup, $contract) } finally { $warmup.Dispose() }
$stream = New-ProbeWorkbook -Kind NamedRange -LastRow $NamedRangeEndRow
try {
    $watch = [Diagnostics.Stopwatch]::new()
    $before = [GC]::GetAllocatedBytesForCurrentThread()
    $watch.Start()
    $parsed = [TemplateFrame.Excel.ExcelTemplateParser]::new().Parse($stream, $contract)
    $watch.Stop()
    $allocated = [GC]::GetAllocatedBytesForCurrentThread() - $before
    $rangeResult = [ordered]@{
        namedRangeEndRow = $NamedRangeEndRow; physicalDataRows = 1; inputBytes = $stream.Length
        rowsReturned = $parsed.Tables['Items'].Count; elapsedMilliseconds = $watch.ElapsedMilliseconds
        cumulativeThreadAllocatedBytes = $allocated
    }
    $parsed = $null
} finally { $stream.Dispose() }

$stream = New-ProbeWorkbook -Kind Dtd
try {
    $null = [TemplateFrame.Excel.Simple.SimpleExcel]::Read($stream, $null)
    $dtdResult = [ordered]@{ rejected = $false; error = $null }
} catch {
    $messages = [Collections.Generic.List[string]]::new()
    $exception = $_.Exception
    while ($null -ne $exception) { $messages.Add($exception.Message); $exception = $exception.InnerException }
    $dtdResult = [ordered]@{
        rejected = $true
        prohibitedDtdMessageObserved = [bool]($messages -match 'DTD is prohibited')
        error = $messages.ToArray()
    }
} finally { $stream.Dispose() }

$stream = New-ProbeWorkbook -Kind ActiveContent
$filled = $null
$archive = $null
try {
    $emptyContract = [TemplateFrame.Contract.TemplateContract]::new()
    $validation = [TemplateFrame.Excel.ExcelTemplateValidator]::new().Validate($stream, $emptyContract)
    $stream.Position = 0
    $filled = [TemplateFrame.Excel.ExcelTemplateFiller]::new().Fill($stream, $emptyContract, [TemplateFrame.Data.FillData]::new())
    $archive = [IO.Compression.ZipArchive]::new($filled.Output, [IO.Compression.ZipArchiveMode]::Read, $true)
    $relationEntry = $archive.GetEntry('xl/worksheets/_rels/sheet1.xml.rels')
    $relationText = ''
    if ($null -ne $relationEntry) {
        $reader = [IO.StreamReader]::new($relationEntry.Open())
        try { $relationText = $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    $activeContentResult = [ordered]@{
        validationIssues = $validation.Issues.Count
        inertVbaMarkerPartPreserved = ($null -ne $archive.GetEntry('xl/vbaProject.bin'))
        externalHyperlinkPreserved = $relationText.Contains('https://example.invalid/safe-inert-probe')
        executedCodeOrAccessedLink = $false
        interpretation = 'Content preservation only; this does not demonstrate macro execution or SSRF.'
    }
} finally {
    if ($null -ne $archive) { $archive.Dispose() }
    if ($null -ne $filled) { $filled.Output.Dispose() }
    $stream.Dispose()
}

$results = [ordered]@{
    auditDate = '2026-10-22'
    executedAtUtc = [DateTime]::UtcNow.ToString('o')
    repositoryRoot = $repository
    configuration = $Configuration
    targetFrameworkOfLoadedArtifacts = 'net8.0'
    actualExecutionRuntime = [Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
    powerShellVersion = $PSVersionTable.PSVersion.ToString()
    assemblyManifest = $assemblyManifest
    memoryMetric = 'Cumulative current-thread managed allocation; NOT peak memory, working set, retained heap, or native memory.'
    caveats = @(
        'Reads prebuilt DLLs; does not build the repository. Compare assembly manifest hashes/timestamps with intended build.',
        'Assembly.LoadFrom may bind an already-loaded PowerShell dependency. path/sha256/diskAssemblyVersion describe the file on disk; loadedAssemblyLocation/loadedAssemblyVersion describe the actual binding. assemblyVersion is the legacy loaded-version field.',
        'Synthetic bounded probes only, not a stress test or complete security test.',
        'No int.MaxValue, archive bomb, executable macro, external network request, or disk extraction probe.',
        'Elapsed time and cumulative allocations are host-dependent observations, not performance guarantees.',
        'This does not validate net462 or a .NET 8 host merely because net8.0 target artifacts were loaded.'
    )
    simpleSparseRows = $simpleResult
    excelNamedRange = $rangeResult
    dtd = $dtdResult
    defaultMaxCharactersInPart = [DocumentFormat.OpenXml.Packaging.OpenSettings]::new().MaxCharactersInPart
    activeContentPreservation = $activeContentResult
}
$json = $results | ConvertTo-Json -Depth 10
[IO.File]::WriteAllText($ResultsPath, $json, [Text.UTF8Encoding]::new($false))
Write-Output $json