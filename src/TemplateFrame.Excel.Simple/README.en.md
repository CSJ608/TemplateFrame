# TemplateFrame.Excel.Simple

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Excel.Simple.svg)](https://www.nuget.org/packages/TemplateFrame.Excel.Simple)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Excel.Simple)](https://www.nuget.org/packages/TemplateFrame.Excel.Simple)

> [中文](README.md) · English

The **simplified Excel plugin** for TemplateFrame: table import/export limited to "header row + data rows".

Most Excel import/export is exactly "a header row, then column after column of data". For that simple shape you don't need
[TemplateFrame.Excel](../TemplateFrame.Excel/README.md) with merged cells / images / layout —
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
using var stream = File.Create("items.xlsx");
SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "Materials" });

// Import (locates the header via the TF_Table named range first; falls back to "first non-empty row with 2+ cells" when absent/misplaced)
using var input = File.OpenRead("items.xlsx");
var loaded = SimpleExcel.Read(input); // Headers + Rows (string / bool / DateTime / double / null)
```

- Cell values support: `string` / `bool` / `DateTime` (written as a date serial + `yyyy-mm-dd`) / numbers / `null`.
- **Named-range location**: `Write` stores the table area as a named range (default `TF_Table` → `'Materials'!$A$1:$C$3`; customize with `TableName`, place with `StartCell`); `Read` locates the header through it first, falling back to "first non-empty row with 2+ cells" when the range is missing **or its header row is empty (misplaced range)** (skipping title/decoration rows with only one non-empty cell).
- **Data-area tolerance**: data rows always extend to the last worksheet row (all-empty rows skipped) — when the named range covers only the header, or a user manually appends data below it in Excel, nothing is silently dropped. Note: non-empty content below the range (e.g. a second table) will be read in as well.
- Compatible with common external files: shared-string headers (the Excel/WPS default) resolve to real text; rich-text cells (partially bold/colored) concatenate all run fragments; rows missing the `RowIndex(r)` attribute are inferred by document order (the extreme case of cells missing `r` references is unsupported — Excel/WPS always write cell references, so real files don't hit it).
- Numbers come back as `double`, date-formatted cells as `DateTime`; all-empty rows skipped, missing columns padded with null.
- No page setup / merged cells / images — keeping the minimal "simple table" shape.

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
var service = new MaterialsTemplateService();
using var template = service.BuildTemplate();          // header-only
var validation = service.Validate(template);           // header ↔ contract columns (missing required column = Error / extra column = Warning)
using var filled = service.Fill(data);                 // typed data → xlsx (header + data rows)
var parsed = service.Parse(filled);                    // xlsx → typed MaterialsData
```

- **Contract shape**: only a **single `TableElement`** is supported (columns = headers); scalar/image elements or multiple tables throw a clear error (that's `TemplateFrame.Excel` territory).
- **Column location (graded fallback)**: read/validate locate columns first via **per-column defined names** (`TF_<TableName>_<ColumnKey>` → header cell, auto-generated when the framework writes the file) — **parsing is decoupled from the header language (language-independent)**; when defined names are unavailable it falls back to header text matching (`DisplayName` → `Key`). Extra columns are ignored, missing columns padded with null; `Validate` reports `Missing` (Error) for required columns, `Warning` for missing optional columns and extra columns, and `Ambiguous` (Error) for duplicated column defined names.
- **Localized headers**: `SimpleExcelContract.Write(..., culture, localizer)` or `service.Fill(data, options, culture, localizer)` write localized headers (localization key = column Key; unregistered keys fall back to `DisplayName`/`Key`); parsing stays language-independent (defined-name location).
- **Low-level API**: you can also use `SimpleExcelContract.Write / Read / Validate` (based on `FillData`) directly with the base package's `DataPathMapper` for custom mapping.
- **Backward compatible**: the original `SimpleExcel.Write / Read` (`SimpleExcelTable`) is unchanged.

## Root collection: fill / parse List<T> directly

If the scenario data is just a list (no wrapper object needed), declare `TData` as a collection type and leave the table's `DataPath` empty — row data is taken from the root object itself:

```csharp
public sealed class MaterialListService : SimpleExcelTemplateService<List<MaterialLine>>
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
                    // DataPath left empty = root collection: TData (List<MaterialLine>) is itself the row collection
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

var service = new MaterialListService();
using var filled = service.Fill(
[
    new MaterialLine { Code = "AL-6063", Name = "Aluminum profile 6063-T5", Qty = 120.5m },
    new MaterialLine { Code = "SS-M8", Name = "Stainless bolt M8x30", Qty = 500m },
]);
var parsed = service.Parse(filled);      // directly yields List<MaterialLine>
```

- **Supported root collection types**: `List<T>` / `IReadOnlyList<T>` / `IEnumerable<T>` / arrays `T[]` (`Parse` returns what was declared; interface collections are backed by `List<T>`).
- With a root collection the table `DataPath` **must be empty** (declaring one throws a clear error); column `DataPath` still points at row-element properties.
- The wrapper-object style (`MaterialsData.Items`) and the low-level `SimpleExcelTable` API are unchanged — fully backward compatible.
- i18n works like the wrapper version: `Fill(..., culture, localizer)` writes localized headers while defined-name parsing stays language-independent (see the root-collection section of `samples/TemplateFrame.Demo.Excel.Simple.I18n`).

## Performance and dependencies

- Measured on an ordinary dev machine (scales linearly with rows): write / read of 1,000 rows ~30ms, 10,000 rows ~0.3–0.5s; contract-path read of 10,000 rows ~0.6–0.9s.
- Snapshots in `docs/PERFORMANCE.md`; benchmark project `test/TemplateFrame.Benchmarks` (reproducible with `dotnet run -c Release`).
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
