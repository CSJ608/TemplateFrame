# TemplateFrame.Excel

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Excel.svg)](https://www.nuget.org/packages/TemplateFrame.Excel)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Excel)](https://www.nuget.org/packages/TemplateFrame.Excel)

> [中文](README.md) · English

The **MS Excel plugin** for TemplateFrame: translates the base package's "contract + data shape" into `.xlsx`.
Built on **named ranges (defined names)** for the full **generate → locate → fill → parse → validate** pipeline, using
DocumentFormat.OpenXml directly (same family as the Word plugin, no new third-party dependencies).

## Design conventions

- **No page setup**: Excel is a "regular grid" layout (unlike Word's paper/orientation/margins),
  so the Builder has no SetPageSetup; width follows the body column count, layout uses merged cells (the demo uses a 3×9 grid header).
- **Word wrap**: `TextFormat.WrapText = true` (enabled for table headers/cells and header values, long text wraps instead of overflowing).
  **Row height**: `SetRowHeight(row, pt)` writes customHeight (together with sheetViews, Excel stops recalculating row heights).
- **For simple tables use TemplateFrame.Excel.Simple**: most import/export is just "header row + data rows"
  without named ranges/merges/images — the separate plugin TemplateFrame.Excel.Simple is more direct.

## Core components

| Component | Responsibility |
|---|---|
| `ExcelTemplateBuilder` | Assembles an .xlsx with named ranges: column widths, row heights, cell formats (incl. word wrap), merged cells, tables (header + sample row), images (cell-anchored + offset) |
| `ExcelNamedRangeLocator` | Locates by named range (`TF_` prefix, workbook-unique): scalar `TF_<Key>` → cell; table column `TF_<TableKey>_<ColumnKey>` → sample row |
| `ExcelTemplateFiller` | Fill: text writes typed values + number formats (dates stored as serial numbers), images swap part + relationship (size inherited), table rows cloned then column ranges re-pointed + elements below shifted down as a block; soft validation before filling |
| `ExcelTemplateParser` | Parse: reads a filled template back into `FillData` per the contract (text converted by ValueType, multi-row tables, image bytes) |
| `ExcelTemplateValidator` | Validate: Missing / WrongType / Ambiguous / Extra (missing optional fields only warn) |

## Location mechanism (named ranges)

Excel has no content controls (SDT), so **named ranges** take over tag-based location:

- Scalar elements: `TF_<Key>` → single cell (e.g. `TF_OrderNo` → `'Delivery'!$B$2`), workbook-unique;
- Tables: each column `TF_<TableKey>_<ColumnKey>` points at the **sample row** cell; during fill the sample row becomes data row 1,
  rows 2..N are cloned, each column range is **re-pointed to the whole data block** (e.g. `$C$5:$C$9`), and named ranges / merged ranges below the table are **shifted down (N-1) rows as a block**;
- Parsing an unfilled template reads placeholder text from the sample row (default zh "待填充" / en "To be filled", generated per language; Parse normalizes known placeholders to null).
- **i18n keys**: `AddTextKey(cellAddress, key, format?)` / `AddTableKeys(key, columnKeys, format?, startCell?)` resolve layout text / headers by language (key methods vs literal methods; each column range `TF_<TableKey>_<ColumnKey>` still uses the column Key, so parsing is independent of header language).

## Quick start

Your scenario service declares the plugin builder type; `BuildInitialTemplate()` takes no arguments and composes directly with the `Builder` instance:

```csharp
public sealed class DeliveryOrderExcelTemplateService : TemplateService<DeliveryOrderData, ExcelTemplateBuilder>
{
    public DeliveryOrderExcelTemplateService() : base(new ExcelTemplateEngine()) { }

    protected override TemplateContract DefineContract() => /* element list (shared with the Word version) */;

    protected override void BuildInitialTemplate()
    {
        Builder.SetSheetName("Delivery");
        // No page setup: Excel lays out via grid + merged cells (3×9 header grid, see the demo)
        Builder.MergeCells("A1:B3"); // LOGO area
        Builder.MergeCells("C1:G3"); // title area
        Builder.AddText("C1", "DELIVERY ORDER", new TextFormat { FontName = "SimHei", SizePt = 16, Bold = true, Alignment = TextAlignment.Center });
        Builder.AddElement("OrderNo", "B2");
        Builder.AddTable("Lines", ["No.", "Code", "Material", "Unit", "Planned", "Received", "Batch", "Supplier Batch", "Warehouse"],
            new TableFormat { HeaderFormat = ..., CellFormat = ..., Bordered = true, ColumnWidthsCm = [...] }, "A6");
        Builder.AddImage("Logo", "H2", 0.8, 0.8);
    }

    protected override FillData MapToData(DeliveryOrderData data) => /* manual mapping */;
    protected override DeliveryOrderData MapFromData(FillData data) => /* manual reverse mapping */;
}
```

## Fill behavior notes

- **Text**: writes **typed values + number formats** (DateTime stored as an OADate serial + date format; decimal/int as numbers; bool as 0/1),
  preserving the cell's existing font/borders/alignment; null writes empty.
- **Images**: locates the drawing by anchor cell, swaps the image part + relationship, updates `r:embed`; size/position inherit the placeholder.
- **Table rows**: the sample row becomes data row 1, rows 2..N are deep-copied (row indices and cell references rewritten), values filled per row;
  after cloning the column ranges are re-pointed to the data block and named ranges / merged ranges below the table shift down as a block.
- **Soft validation** (Validate runs before filling): `Drifted`/`Extra` only record warnings and continue; missing required elements follow the policy (throw by default, configurable via `SkipAndWarn`);
  `WrongType`/`Ambiguous`/`Invalid` are hard errors.
- **Warning outlet**: `ExcelTemplateFiller.Fill` returns a `TemplateFillResult` (output stream + Warnings); the engine/service layer offers `FillDetailed` (`ITemplateEngine.FillDetailed` / `TemplateService<TData, TBuilder>.FillDetailed`) for the same soft-validation warnings, while `Fill` keeps returning only the output stream.

## Parse behavior notes

- Text converts by `TextElement.ValueType` (string/decimal/int/DateTime/bool; dates restored from serial numbers); tables read back row by row along column range extents (columns aligned by row index); images read back as bytes.
- Known placeholders in unfilled templates (default zh "待填充" / en "To be filled") normalize to **null**.

## Dependencies and tests

- Target frameworks `netstandard2.0 / net462 / net8.0` (NuGet picks per runtime automatically).
- Depends on `DocumentFormat.OpenXml` (3.3.x, same as the Word plugin).
- Performance (measured on an ordinary dev machine, scales linearly with rows): 1k-row detail fill ~60ms, parse ~115ms, build ~1ms; snapshots in `docs/PERFORMANCE.md`, benchmark project `test/TemplateFrame.Benchmarks`.
- Tests in `test/TemplateFrame.Excel.Tests`: generate → validate → fill → parse → assert (including named-range inventories, typed values,
  range re-pointing after row cloning, elements below shifted down, image replacement, unfilled placeholders and other edge cases).

## Full example

See the **Excel delivery order** in `samples/TemplateFrame.Demo.Excel` (reuses the delivery-order data, 3×9 grid header / 9-column detail):

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel
```

Design doc: `docs/DESIGN.md`; usage guide: the repository root `README.md` (Chinese) / `README.en.md` (English).
