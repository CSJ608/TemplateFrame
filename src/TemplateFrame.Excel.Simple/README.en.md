# TemplateFrame.Excel.Simple

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Excel.Simple.svg)](https://www.nuget.org/packages/TemplateFrame.Excel.Simple)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Excel.Simple)](https://www.nuget.org/packages/TemplateFrame.Excel.Simple)

> [中文](https://github.com/CSJ608/TemplateFrame/blob/main/src/TemplateFrame.Excel.Simple/README.md) · English

The **simplified Excel plugin** for TemplateFrame: table import/export limited to "header row + data rows".

Most Excel import/export is exactly "a header row, then column after column of data". For that simple shape you don't need
[TemplateFrame.Excel](https://github.com/CSJ608/TemplateFrame/blob/main/src/TemplateFrame.Excel/README.en.md) with merged cells / images / layout —
the two plugins split the two different needs:

| Plugin | Positioning | Capabilities |
|---|---|---|
| `TemplateFrame.Excel` | Flexible layout (documents / complex tables) | Named-range location, merges, images, table cloning, Validate/Fill/Parse |
| `TemplateFrame.Excel.Simple` | Simple tables (header + data rows) | Write / Read; a named range marks the table location (default `TF_Table`); no page setup, no merges, no images |

## Usage

```csharp
using TemplateFrame.Excel.Simple;

// Export (writes from A1 by default; the named range TF_Table marks the table area; StartCell / TableName customize it)
var table = new SimpleExcelTable
{
    Headers = ["Code", "Material", "Qty"],
    Rows =
    [
        ["AL-6063", "Aluminum profile 6063-T5", 120m],
        ["SS-M8", "Stainless bolt M8x30", 500m],
    ],
};
using (var stream = File.Create("items.xlsx"))
{
    SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "Materials" });
}

// Import (locates the header via the TF_Table named range first; falls back to "first non-empty row with 2+ cells" when absent/misplaced)
using var input = File.OpenRead("items.xlsx");
var loaded = SimpleExcel.Read(input); // Headers + Rows (string / bool / DateTime / double / null)
```

- `Write/Read` handles header + data rows, with no merges, images or page setup. Read returns numbers as `double`, date-formatted values as `DateTime`, and missing cells as null; writing decimal/long without double conversion does not imply equal read-back precision.
- Read uses the named range, with header fallback when it is absent or empty. It reads to the worksheet's last row and skips empty rows; unrelated content below the table can also be included. Full [location, fallback and value rules](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences) are in DESIGN (Chinese).

## Contract + strongly-typed service

Simple tables can also join the TemplateFrame contract system, so `service.Parse` yields strongly-typed data just like Word:

```csharp
using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;

public sealed record MaterialLine
{
    public string Code { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public decimal Qty { get; init; }
}

public sealed record MaterialsData
{
    public IReadOnlyList<MaterialLine> Items { get; init; } = [];
}

public sealed class MaterialsTemplateService : SimpleExcelTemplateService<MaterialsData>
{
    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "Materials",
            Version = "1.0",
            Elements =
            [
                new TableElement
                {
                    Key = "Materials",
                    DisplayName = "Materials",
                    DataPath = "Items",                      // table → collection property
                    Columns =
                    [
                        new TextElement { Key = "Code", DisplayName = "Code", DataPath = "Code", Required = true },
                        new TextElement { Key = "Name", DisplayName = "Name", DataPath = "Name", Required = true },
                        new TextElement { Key = "Qty", DisplayName = "Qty", DataPath = "Qty", ValueType = typeof(decimal) },
                    ],
                },
            ],
        };
}

// Usage: contract → strong types (tables and columns with DataPath map automatically; no hand-written MapToData / MapFromData)
var data = new MaterialsData { Items = [new MaterialLine { Code = "M001", Name = "Bolt", Qty = 10m }] };
var service = new MaterialsTemplateService();
using var template = service.BuildTemplate();          // header-only
var validation = service.Validate(template);           // header ↔ contract columns (missing required column = Error / extra column = Warning)
using var filled = service.Fill(data);                 // typed data → xlsx (header + data rows)
var parsed = service.Parse(filled);                    // xlsx → typed MaterialsData
```

- A contract contains one `TableElement`. Use `Validate` to check required/missing, extra or ambiguous columns. Valid per-column defined names make parsing independent of header language. In text fallback, both `Validate` and `Read` match the non-empty trimmed DisplayName first, then the trimmed Key only if DisplayName is absent from the headers, using case-sensitive Ordinal comparison. DisplayName wins when both appear, regardless of physical order. If contract columns select the same header, Read assigns it to the first declared column; repeated physical headers still use the last value. Missing columns do not fabricate values; no new text ambiguity diagnostics are added. See [matching rules and boundaries](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences) (Chinese).
- `SimpleExcelTemplateService` is independent of the Word/Excel service: no Builder/Engine, `FillDetailed` or `ParseDetailed`, and no Builder generation lock. Its default Parse mapping is strict. See [Simple rules](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences) and [mapping/cache rules](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md).
- Use `culture` / `localizer` on Fill for localized headers. Lower-level APIs are `SimpleExcelContract.Write / Read / Validate` with `FillData`.

## Root collections

For list-only data, inherit `SimpleExcelTemplateService<List<MaterialLine>>`. Keep the column declarations above, leave the table DataPath empty, and pass the list to Fill; Parse returns the declared collection type. Supported types and constraints are in [DESIGN §3.3](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md). A runnable localized example is [Simple.I18n](https://github.com/CSJ608/TemplateFrame/tree/main/samples/TemplateFrame.Demo.Excel.Simple.I18n).

## Performance and dependencies

- Historical snapshot (2026-08-24; sample-specific, with no guarantee of linear scaling or current-version timings): write / read of 1,000 rows ~30ms, 10,000 rows ~0.3–0.5s; contract-path read of 10,000 rows ~0.6–0.9s.
- Snapshots in [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md); benchmark project [benchmarks](https://github.com/CSJ608/TemplateFrame/blob/main/test/TemplateFrame.Benchmarks/README.md) (see that page for reproduction commands).
- Target frameworks `netstandard2.0 / net462 / net8.0` (NuGet picks per runtime automatically); depends on `DocumentFormat.OpenXml` (3.3.x).

## Demo

`samples/TemplateFrame.Demo.Excel.Simple` in the repository provides the **material master data** example (template → fill → re-parse full loop; headers: code / name / unit / package spec / model):

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel.Simple
```

Outputs go to the system temp directory `%TEMP%\TemplateFrame.Demo.Excel.Simple` by default:
- `Excel-Simple-Materials-template.xlsx`: the **template** (header-only, defines the column structure)
- `Excel-Simple-Materials-filled.xlsx`: the **filled** file (header + material data rows)
- The console prints the **re-parse** result (read the filled file → `SimpleExcel.Read` → print headers and each row)
