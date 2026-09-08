# TemplateFrame.Word

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Word.svg)](https://www.nuget.org/packages/TemplateFrame.Word)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Word)](https://www.nuget.org/packages/TemplateFrame.Word)

> [中文](https://github.com/CSJ608/TemplateFrame/blob/main/src/TemplateFrame.Word/README.md) · English

The **MS Word plugin** for TemplateFrame: translates the base package's "contract + data shape" into `.docx`.
Built on content controls (SDT / Structured Document Tags) for the full **generate → locate → fill → parse → validate** pipeline.
Only Microsoft Office `.docx` is supported (for WPS scope see [DESIGN](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md)).

## Core components

| Component | Responsibility |
|---|---|
| `WordTemplateBuilder` | Assembles an SDT-tagged .docx: page setup, header/footer, layout tables, detail tables, text/image elements, page number fields |
| `SdtLocator` | Locates content controls by tag (body/header/footer; scalar tags are unique, table columns are located per row) |
| `WordTemplateFiller` | Fill: text (preserving run formatting), images (swap package part + relationship), table row cloning (re-issuing unique w:id); soft validation before filling |
| `WordTemplateParser` | Parse: reads a filled template back into `FillData` per the contract (text converted by ValueType, multi-row tables, image bytes) |
| `WordTemplateValidator` | Validate: Missing / WrongType / Ambiguous / Extra (missing optional fields only warn) |

## Quick start

Your scenario service declares the plugin builder type; `BuildInitialTemplate()` takes no arguments and composes directly with the `Builder` instance. This fragment omits the DTO, contract and mapping; see the runnable example below:

```csharp
public sealed class DeliveryOrderTemplateService : TemplateService<DeliveryOrderData, WordTemplateBuilder>
{
    public DeliveryOrderTemplateService() : base(new WordTemplateEngine()) { }

    protected override TemplateContract DefineContract() => /* element list */;

    protected override void BuildInitialTemplate()
    {
        Builder.SetPageSetup(new PageSetup { Size = PageSize.A5, Orientation = PageOrientation.Landscape });
        Builder.AddHeader(BuildHeader);   // header (LOGO / title / QR code + page number)
        Builder.AddFooter(BuildFooter);   // footer (date / receiver / page number)
        Builder.AddTable("Lines", ["No.", "Material", "Qty", "Unit"],
            new TableFormat
            {
                HeaderFormat = new TextFormat { FontName = "SimSun", SizePt = 12, Bold = true, Alignment = TextAlignment.Center },
                CellFormat = new TextFormat { FontName = "SimSun", SizePt = 12, Alignment = TextAlignment.Center },
                Alignment = TextAlignment.Center,
                ColumnWidthsCm = [1.2, 6.0, 2.5, 2.0],
            });
    }

    protected override FillData MapToData(DeliveryOrderData data) => /* manual mapping */;
    protected override DeliveryOrderData MapFromData(FillData data) => /* manual reverse mapping */;
}
```

## WordTemplateBuilder capabilities (the typed methods)

- **Page**: `SetPageSetup(PageSetup)` — A4/A5, portrait/landscape, millimeter margins
- **Header/footer**: `AddHeader(Action<WordTemplateBuilder>)` / `AddFooter(...)` — same capabilities as the body
- **Layout table**: `AddLayoutTable(rows, cols, TableFormat?)` + `AddCell(compose, columnSpan)` — header "left/center/right / equal split / four cells" (gridSpan column spanning)
- **Text**: `AddParagraph(text[, style|TextFormat])` / `AddText` / `AddElement(key[, TextFormat])` (element = content control; placeholder text follows the language: default zh "待填充" / en "To be filled", resolved via `ITemplateLocalizer`, overridable)
- **Table**: `AddTable(key, columns, TableFormat?, headerStyle?)` — header + sample row (one SDT per cell); `TableFormat` supports header/cell fonts, borders on/off, table alignment, column widths (cm), vertical alignment
- **Image**: `AddImage(key, placeholderPath?, widthInches?, heightInches?)` — placeholder image wrapped in an SDT; filling swaps in `byte[]`
- **Page number**: `AddPageNumber(pattern? = null, TextFormat?)` — PAGE/NUMPAGES fields; a null pattern picks the language default (zh "第{page}页，总{total}页" / en "Page {page} of {total}")
- `TextFormat`: `FontName` / `SizePt` / `Bold` / `Alignment` / `Underline`

## Format notes

- SDT tags locate fields in the body, headers and footers. Text updates preserve nested controls; images use relationships in their owning part. Cloned SDTs receive unique `w:id` values.
- Fill from the original unfilled template each time, including before/after-receipt outputs. Zero rows clear sample placeholders while preserving the header and a blank row.
- Parse reads text, table rows and image bytes. Known placeholders normalize to null. Word conversion messages use data row numbers; Excel messages use worksheet row numbers. Structured `DataRowNumber` is one-based in both plugins.

Use `FillDetailed` / `ParseDetailed` for warnings. Engine failures retain raw text; default typed mapping retains property defaults. Full [validation, mapping and diagnostic rules](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#parse-diagnostics) and [service lifetime rules](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#service-lifetime) are maintained in DESIGN (Chinese). See [format differences](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences) for the separate Simple API.

## Dependencies and tests

- Target frameworks `netstandard2.0 / net462 / net8.0` (NuGet picks per runtime automatically).
- Depends on `DocumentFormat.OpenXml` (3.3.x).
- Tests in `test/TemplateFrame.Word.Tests`: generate → validate → fill → parse → assert (including header/footer, multi-table, batch, spanning layout, header image part ownership edge cases).
- Historical snapshot (2026-08-24; sample-specific, with no guarantee of linear scaling or current-version timings): 1k-row detail fill ~150ms, parse ~125ms, build <1ms; snapshots in [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md), benchmark project [benchmarks](https://github.com/CSJ608/TemplateFrame/blob/main/test/TemplateFrame.Benchmarks/README.md).

## Full example

See the **delivery order** in `samples/TemplateFrame.Demo.Word` (two-tier header + 9-column detail + two-line footer + before/after-receipt double fill):

```bash
dotnet run --project samples/TemplateFrame.Demo.Word
```

Design: [DESIGN](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md) (Chinese); getting started: [English README](https://github.com/CSJ608/TemplateFrame/blob/main/README.en.md) / [中文 README](https://github.com/CSJ608/TemplateFrame/blob/main/README.md).
