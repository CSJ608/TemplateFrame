using System.Diagnostics;
using System.IO.Compression;
using System.IO.Packaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel;
using TemplateFrame.Excel.Simple;

const int sparseRow = 100_000;
const int namedRangeEndRow = 10_000;
const string repositoryRoot = @"C:\work\OpenCode\TemplateFrame";
var outputPath = args.Length > 0 ? args[0] : Path.Combine(Path.GetTempPath(), "TemplateFrame-security-audit-20261022", "input-probes-net8-results.json");
if (Environment.Version.Major != 8) throw new InvalidOperationException($"Expected .NET 8 host; got {RuntimeInformation.FrameworkDescription}");
if (typeof(Package).Assembly.GetName().Version?.Major != 8) throw new InvalidOperationException($"Expected System.IO.Packaging 8; got {typeof(Package).Assembly.FullName}");

using (var warmup = NewWorkbook("Simple", 2)) _ = SimpleExcel.Read(warmup);
object simpleResult;
using (var stream = NewWorkbook("Simple", sparseRow))
{
    var watch = new Stopwatch();
    var before = GC.GetAllocatedBytesForCurrentThread();
    watch.Start();
    var parsed = SimpleExcel.Read(stream);
    watch.Stop();
    var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
    simpleResult = new { lastRowIndex = sparseRow, physicalDataRows = 1, inputBytes = stream.Length, rowsReturned = parsed.Rows.Count, elapsedMilliseconds = watch.ElapsedMilliseconds, cumulativeThreadAllocatedBytes = allocated };
}

var contract = new TemplateContract { Elements = [new TableElement { Key = "Items", Columns = [new TextElement { Key = "Value" }] }] };
using (var warmup = NewWorkbook("NamedRange", 2)) _ = new ExcelTemplateParser().Parse(warmup, contract);
object rangeResult;
using (var stream = NewWorkbook("NamedRange", namedRangeEndRow))
{
    var watch = new Stopwatch();
    var before = GC.GetAllocatedBytesForCurrentThread();
    watch.Start();
    var parsed = new ExcelTemplateParser().Parse(stream, contract);
    watch.Stop();
    var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
    rangeResult = new { namedRangeEndRow, physicalDataRows = 1, inputBytes = stream.Length, rowsReturned = parsed.Tables["Items"].Count, elapsedMilliseconds = watch.ElapsedMilliseconds, cumulativeThreadAllocatedBytes = allocated };
}

object dtdResult;
using (var stream = NewWorkbook("Dtd", 2))
{
    try { _ = SimpleExcel.Read(stream); dtdResult = new { rejected = false }; }
    catch (Exception ex)
    {
        var messages = new List<string>();
        for (Exception? current = ex; current != null; current = current.InnerException) messages.Add(current.Message);
        dtdResult = new { rejected = true, prohibitedDtdMessageObserved = messages.Any(m => m.Contains("DTD is prohibited", StringComparison.Ordinal)), errors = messages };
    }
}

object activeContentResult;
using (var stream = NewWorkbook("ActiveContent", 2))
{
    var emptyContract = new TemplateContract();
    var validation = new ExcelTemplateValidator().Validate(stream, emptyContract);
    stream.Position = 0;
    var filled = new ExcelTemplateFiller().Fill(stream, emptyContract, new FillData());
    using var output = filled.Output;
    using var archive = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true);
    var relationEntry = archive.GetEntry("xl/worksheets/_rels/sheet1.xml.rels");
    string relationText = "";
    if (relationEntry != null) { using var reader = new StreamReader(relationEntry.Open()); relationText = reader.ReadToEnd(); }
    activeContentResult = new { validationIssues = validation.Issues.Count, inertVbaMarkerPartPreserved = archive.GetEntry("xl/vbaProject.bin") != null, externalHyperlinkPreserved = relationText.Contains("https://example.invalid/safe-inert-probe", StringComparison.Ordinal), executedCodeOrAccessedLink = false };
}

var assemblies = new[] { typeof(Package).Assembly, typeof(OpenSettings).Assembly, typeof(SpreadsheetDocument).Assembly, typeof(TemplateContract).Assembly, typeof(ExcelTemplateParser).Assembly, typeof(SimpleExcel).Assembly };
var manifest = assemblies.Distinct().Select(assembly =>
{
    var location = assembly.Location;
    var hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(location)));
    var sourceDirectory = assembly.GetName().Name == "TemplateFrame.Excel.Simple" ? "TemplateFrame.Excel.Simple.Tests" : "TemplateFrame.Excel.Tests";
    var sourcePath = Path.Combine(repositoryRoot, "test", sourceDirectory, "bin", "Release", "net8.0", Path.GetFileName(location));
    var sourceHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(sourcePath)));
    if (hash != sourceHash) throw new InvalidOperationException($"Loaded assembly hash differs from freshly built source artifact: {location}");
    return new { name = assembly.GetName().Name, assemblyVersion = assembly.GetName().Version?.ToString(), fileVersion = FileVersionInfo.GetVersionInfo(location).FileVersion, loadedLocation = location, sha256 = hash, sourceArtifact = sourcePath, sourceSha256 = sourceHash, sourceLastWriteTimeUtc = File.GetLastWriteTimeUtc(sourcePath), loadedHashMatchesSource = true };
}).ToArray();

var packagingLocation = typeof(Package).Assembly.Location;
var expectedPackagePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".nuget", "packages", "system.io.packaging", "8.0.1", "lib", "net8.0", "System.IO.Packaging.dll");
var actualPackagingHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(packagingLocation)));
var expectedPackagingHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(expectedPackagePath)));
if (actualPackagingHash != expectedPackagingHash) throw new InvalidOperationException("Loaded Packaging assembly is not identical to NuGet System.IO.Packaging 8.0.1/net8.0.");

var results = new
{
    auditDate = "2026-10-22", executedAtUtc = DateTime.UtcNow, repositoryRoot,
    probeSource = "Independent net8.0 console using direct references to rebuilt workspace Release DLLs; no PowerShell runtime assembly binding.",
    actualExecutionRuntime = RuntimeInformation.FrameworkDescription, runtimeVersion = Environment.Version.ToString(), processPath = Environment.ProcessPath,
    assemblyManifest = manifest,
    packagingPackageVerification = new { expectedNuGetPackage = "System.IO.Packaging/8.0.1", expectedPackagePath, actualPackagingHash, expectedPackagingHash, exactBinaryMatch = true },
    memoryMetric = "Cumulative current-thread managed allocation; NOT peak memory, working set, retained heap, or native memory.",
    simpleSparseRows = simpleResult, excelNamedRange = rangeResult, dtd = dtdResult,
    defaultMaxCharactersInPart = new OpenSettings().MaxCharactersInPart, activeContentPreservation = activeContentResult,
    limitations = new[] { "Bounded synthetic probes only; no stress test.", "No executable macro, external request, filesystem extraction, or archive bomb.", "Times and allocations are host-dependent observations, not performance guarantees.", "No product source changes; project and build outputs are under TEMP.", "Does not test net462 host behavior." }
};
var json = JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true });
File.WriteAllText(outputPath, json, new UTF8Encoding(false));
Console.WriteLine(json);

static MemoryStream NewWorkbook(string kind, int lastRow)
{
    var stream = new MemoryStream();
    using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
    {
        var workbookType = kind == "ActiveContent" ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        var extraType = kind == "ActiveContent" ? "<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>" : "";
        var extraRelation = kind == "ActiveContent" ? "<Relationship Id=\"rIdVba\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>" : "";
        var definedNames = kind == "NamedRange" ? $"<definedNames><definedName name=\"TF_Items_Value\">Sheet1!$A$2:$A${lastRow}</definedName></definedNames>" : "";
        var worksheet = kind == "NamedRange"
            ? "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>Data</t></is></c></row></sheetData></worksheet>"
            : $"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Header</t></is></c></row><row r=\"{lastRow}\"><c r=\"A{lastRow}\" t=\"inlineStr\"><is><t>Data</t></is></c></row></sheetData></worksheet>";
        if (kind == "Dtd") worksheet = "<!DOCTYPE worksheet [<!ENTITY x \"safe-probe\">]>" + worksheet;
        if (kind == "ActiveContent") worksheet = worksheet.Replace("</worksheet>", "<hyperlinks><hyperlink ref=\"A1\" r:id=\"rIdExternal\"/></hyperlinks></worksheet>");
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = $"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"{workbookType}\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>{extraType}</Types>",
            ["_rels/.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>",
            ["xl/workbook.xml"] = $"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>{definedNames}</workbook>",
            ["xl/_rels/workbook.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>{extraRelation}</Relationships>",
            ["xl/worksheets/sheet1.xml"] = worksheet
        };
        if (kind == "ActiveContent")
        {
            parts["xl/vbaProject.bin"] = "BENIGN_INERT_MARKER_NOT_EXECUTABLE";
            parts["xl/worksheets/_rels/sheet1.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdExternal\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.invalid/safe-inert-probe\" TargetMode=\"External\"/></Relationships>";
        }
        foreach (var entry in parts) { using var writer = new StreamWriter(archive.CreateEntry(entry.Key).Open()); writer.Write(entry.Value); }
    }
    stream.Position = 0;
    return stream;
}