# TemplateFrame.Word

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Word.svg)](https://www.nuget.org/packages/TemplateFrame.Word)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Word)](https://www.nuget.org/packages/TemplateFrame.Word)

> [中文](README.md) · English

The **MS Word plugin** for TemplateFrame: translates the base package's "contract + data shape" into `.docx`.
Built on content controls (SDT / Structured Document Tags) for the full **generate → locate → fill → parse → validate** pipeline.
Only Microsoft Office `.docx` is supported (for WPS compatibility see the design doc §1.4).

## Core components

| Component | Responsibility |
|---|---|
| `WordTemplateBuilder` | Assembles an SDT-tagged .docx: page setup, header/footer, layout tables, detail tables, text/image elements, page number fields |
| `SdtLocator` | Locates content controls by tag (body/header/footer; tags are globally unique) |
| `WordTemplateFiller` | Fill: text (preserving run formatting), images (swap package part + relationship), table row cloning (re-issuing unique w:id); soft validation before filling |
| `WordTemplateParser` | Parse: reads a filled template back into `FillData` per the contract (text converted by ValueType, multi-row tables, image bytes) |
| `WordTemplateValidator` | Validate: Missing / WrongType / Ambiguous / Extra (missing optional fields only warn) |

## Quick start

Your scenario service declares the plugin builder type; `BuildInitialTemplate()` takes no arguments and composes directly with the `Builder` instance:

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
- **Image**: `AddImage(key, placeholder?, widthIn?, heightIn?)` — placeholder image wrapped in an SDT; filling swaps in `byte[]`
- **Page number**: `AddPageNumber(pattern? = null, TextFormat?)` — PAGE/NUMPAGES fields; a null pattern picks the language default (zh "第{page}页，总{total}页" / en "Page {page} of {total}")
- `TextFormat`: `FontName` / `SizePt` / `Bold` / `Alignment` / `Underline`

## Fill behavior notes

- **Text**: replaces the first `w:r/w:t` inside `sdtContent` (preserving run formatting); leading/trailing spaces get `xml:space="preserve"`.
- **Images**: adds an image part + relationship to the package for a new `rId` and replaces the `<a:blip r:embed>` inside the SDT; size/position/wrapping inherit the placeholder; **image parts inside headers/footers belong to the corresponding Header/Footer rels**.
- **Table rows**: deep-copies the sample row N times and fills by tag per row; **every SDT gets a fresh unique `w:id` after cloning**.
- **Soft validation** (Validate runs before filling): `Drifted`/`Extra` only record warnings and continue; missing required elements follow the policy (throw by default, configurable via `MissingElementPolicy.SkipAndWarn`); `WrongType`/`Ambiguous`/`Invalid` are hard errors.
- **Warning outlet**: `WordTemplateFiller.Fill` returns a `TemplateFillResult` (output stream + Warnings); the engine/service layer offers `FillDetailed` (`ITemplateEngine.FillDetailed` / `TemplateService<TData, TBuilder>.FillDetailed`) for the same soft-validation warnings, while `Fill` keeps returning only the output stream.
- **ParseDetailed (2.3.0)**: the import-side counterpart — fields whose conversion fails keep their raw text and are reported as `ConversionFailed` (Warning, table columns carry the data row number) in a `TemplateParseResult`; null still means not filled, `Parse` is unchanged.
- **Before/after receipt**: the same template filled twice — pass `null` for empty fields before receipt (rendered empty), fill them in after.

## Parse behavior notes

- Text converts by `TextElement.ValueType` (string/decimal/int/DateTime/bool); tables read back row by row from the cloned sample-row region; images read back as bytes.
- **Parse normalization**: known placeholders in unfilled templates (default zh "待填充" / en "To be filled", independent of the template language) normalize to **null** (null = not filled, "" = intentionally blank).

## Dependencies and tests

- Target frameworks `netstandard2.0 / net462 / net8.0` (NuGet picks per runtime automatically).
- Depends on `DocumentFormat.OpenXml` (3.3.x).
- Tests in `test/TemplateFrame.Word.Tests`: generate → validate → fill → parse → assert (including header/footer, multi-table, batch, spanning layout, header image part ownership edge cases).
- Performance (measured on an ordinary dev machine, scales linearly with rows): 1k-row detail fill ~150ms, parse ~125ms, build <1ms; snapshots in `docs/PERFORMANCE.md`, benchmark project `test/TemplateFrame.Benchmarks`.

## Full example

See the **delivery order** in `samples/TemplateFrame.Demo.Word` (two-tier header + 9-column detail + two-line footer + before/after-receipt double fill):

```bash
dotnet run --project samples/TemplateFrame.Demo.Word
```

Design doc: `docs/DESIGN.md`; usage guide: the repository root `README.md` (Chinese) / `README.en.md` (English).
