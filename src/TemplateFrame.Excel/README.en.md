# TemplateFrame.Excel

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Excel.svg)](https://www.nuget.org/packages/TemplateFrame.Excel)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Excel)](https://www.nuget.org/packages/TemplateFrame.Excel)

> [中文](https://github.com/CSJ608/TemplateFrame/blob/main/src/TemplateFrame.Excel/README.md) · English

The **MS Excel plugin** for TemplateFrame: translates the base package's "contract + data shape" into `.xlsx`.
Built on **named ranges (defined names)** for the full **generate → locate → fill → parse → validate** pipeline, using
DocumentFormat.OpenXml directly (same family as the Word plugin, no new third-party dependencies).

## Design conventions

- **No page setup**: Excel is a "regular grid" layout (unlike Word's paper/orientation/margins),
  so the Builder has no SetPageSetup; width follows the body column count, layout uses merged cells (the demo uses a 3×9 grid header).
- **Word wrap**: `TextFormat.WrapText = true` (enabled for table headers/cells and header values, long text wraps instead of overflowing).
  **Row height**: `SetRowHeight(row, pt)` writes the row height and customHeight; verify rendering in the target reader.
- **For simple tables use TemplateFrame.Excel.Simple**: most import/export is just "header row + data rows"
  without free layout/merges/images — the separate plugin TemplateFrame.Excel.Simple is more direct.

## Core components

| Component | Responsibility |
|---|---|
| `ExcelTemplateBuilder` | Assembles an .xlsx with named ranges: column widths, row heights, cell formats (incl. word wrap), merged cells, tables (header + sample row), images (cell-anchored + offset) |
| `ExcelNamedRangeLocator` | Locates by named range (`TF_` prefix, workbook-unique): scalar `TF_<Key>` → cell; table column `TF_<TableKey>_<ColumnKey>` → sample row |
| `ExcelTemplateFiller` | Fill: text writes typed values + number formats (dates stored as serial numbers), images swap part + relationship (size inherited), table rows cloned then column ranges re-pointed + elements below shifted down as a block; soft validation before filling |
| `ExcelTemplateParser` | Parse: reads a filled template back into `FillData` per the contract (text converted by ValueType, multi-row tables, image bytes) |
| `ExcelTemplateValidator` | Validate: Missing / WrongType / Ambiguous / Extra (missing optional fields only warn) |

## Location and localization

Named ranges identify scalar cells and table columns; preserve their references when editing. `AddTextKey` / `AddTableKeys` localize layout text and headers while column keys remain unchanged. Full [location and expansion rules](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences) are maintained in DESIGN (Chinese).

## Quick start

Your scenario service declares the plugin builder type; `BuildInitialTemplate()` takes no arguments and composes directly with the `Builder` instance. This fragment omits the DTO, contract and mapping; see the runnable example below:

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
        Builder.AddElement("OrderNo", "B4");
        Builder.AddTable("Lines", ["No.", "Code", "Material", "Unit", "Planned", "Received", "Batch", "Supplier Batch", "Warehouse"],
            new TableFormat { HeaderFormat = ..., CellFormat = ..., Bordered = true, ColumnWidthsCm = [...] }, "A6");
        Builder.AddImage("Logo", "H2", 0.8, 0.8);
    }

    protected override FillData MapToData(DeliveryOrderData data) => /* manual mapping */;
    protected override DeliveryOrderData MapFromData(FillData data) => /* manual reverse mapping */;
}
```

## Format notes

- Cells store typed values and number formats; dates use serial numbers and bool uses 0/1. Images retain placeholder size and position. Table expansion updates column ranges and shifts lower rows, ranges, merges and image anchors.
- Fill from the original unfilled template each time. Zero rows clear sample placeholders while preserving the header and a blank row.
- Parse uses column ranges and contract ValueType; known placeholders normalize to null. Excel conversion messages use absolute worksheet rows, while structured `DataRowNumber` is a one-based data row number.

Use `FillDetailed` / `ParseDetailed` for warnings. Engine failures retain raw text; default typed mapping retains property defaults. Full [validation, mapping and diagnostic rules](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#parse-diagnostics) and [service lifetime rules](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#service-lifetime) are maintained in DESIGN (Chinese). See [format differences](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences) for the separate Simple API.

## Dependencies and tests

- Target frameworks `netstandard2.0 / net462 / net8.0` (NuGet picks per runtime automatically).
- Depends on `DocumentFormat.OpenXml` (3.3.x, same as the Word plugin).
- Historical snapshot (2026-08-24; sample-specific, with no guarantee of linear scaling or current-version timings): 1k-row detail fill ~60ms, parse ~115ms, build ~1ms; snapshots in [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md), benchmark project [benchmarks](https://github.com/CSJ608/TemplateFrame/blob/main/test/TemplateFrame.Benchmarks/README.md).
- Tests in `test/TemplateFrame.Excel.Tests`: generate → validate → fill → parse → assert (including named-range inventories, typed values,
  range re-pointing after row cloning, elements below shifted down, image replacement, unfilled placeholders and other edge cases).

## Full example

See the **Excel delivery order** in `samples/TemplateFrame.Demo.Excel` (reuses the delivery-order data, 3×9 grid header / 9-column detail):

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel
```

Design: [DESIGN](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md) (Chinese); getting started: [English README](https://github.com/CSJ608/TemplateFrame/blob/main/README.en.md) / [中文 README](https://github.com/CSJ608/TemplateFrame/blob/main/README.md).

See the R5 section of repository [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md) for the Excel lookup optimization and three-size comparison. Historical timings do not promise current performance; cumulative managed allocation is not peak memory.
