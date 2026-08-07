# TemplateFrame

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.svg)](https://www.nuget.org/packages/TemplateFrame)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame)](https://www.nuget.org/stats/packages/TemplateFrame?groupby=Version)

> **[中文版](README.md)** · English

A "template ⇄ data" contract engine: declare a template contract (a list of elements) in code, let a business service assemble the initial template with a concrete plugin builder, validate that an uploaded template matches the contract, then fill it with strongly-typed data — or read data back from a filled template.

- **Three-layer architecture**: base package `TemplateFrame` (generic, stable) + plugins `TemplateFrame.Word` (MS Word) / `TemplateFrame.Excel` (MS Excel, flexible layout) / `TemplateFrame.Excel.Simple` (MS Excel, simple tables) + business scene services (strongly-typed, declared inside the business app).
- **Four operations**: `BuildInitialTemplateFile` / `Validate` / `Fill` (strongly-typed) / `Parse` (strongly-typed read-back).
- **Plugin-based**: Word / Excel (flexible layout) / Excel.Simple (header row + data rows) are supported; demos cover both **hand-written mapping** and **DataPath auto-mapping** service styles. WPS Word and label templates are future work.

Design document: [docs/DESIGN.md](docs/DESIGN.md) · Roadmap: [docs/ROADMAP.md](docs/ROADMAP.md) · **Demo guide: [docs/DEMOS.md](docs/DEMOS.md)** · Publishing: [docs/PUBLISHING.md](docs/PUBLISHING.md).

## Core ideas

- **Contract = element list**: `TemplateContract` describes which elements a scene has (`TextElement` / `ImageElement` / `TableElement`); it is serializable and versionable.
- **Templates belong to the business app**: the contract does not produce layout. A business service assembles the initial template with a concrete plugin builder (e.g., `WordTemplateBuilder`) — headings, static text, content controls, tables, image placeholders, headers/footers.
- **Data shape `FillData`**: a plugin-agnostic weakly-typed container (`Values` scalars + `Tables` detail rows). Type conversion happens at the business service boundary — elements declaring `DataPath` are mapped automatically by `DataPathMapper` (iteration 9), or you hand-write `MapToData` / `MapFromData`.
- **Export and import are two directions of the same contract**: `Fill` (template + data → file) and `Parse` (file → data) share the same tag-based location logic.

## Quick start

### 1. Define a business service (declare contract + assemble layout + mapping)

Define a strongly-typed service that inherits `TemplateService<TData, TBuilder>` (TBuilder is the plugin builder type):

```csharp
public sealed class ReceivingOrderTemplateService : TemplateService<ReceivingOrderData, WordTemplateBuilder>
{
    public ReceivingOrderTemplateService() : base(new WordTemplateEngine()) { }

    protected override TemplateContract DefineContract() => new()
    {
        Name = "ReceivingOrder",
        Version = "1.0",
        Elements =
        [
            new TextElement { Key = "OrderNo", DisplayName = "单号", Required = true },
            new TableElement
            {
                Key = "Lines",
                DisplayName = "明细行",
                Columns =
                [
                    new TextElement { Key = "MC", DisplayName = "物料代码" },
                    new TextElement { Key = "Qty", DisplayName = "数量" },
                ],
            },
            new ImageElement { Key = "Logo", DisplayName = "单据图片" },
        ],
    };

    protected override void BuildInitialTemplate()
    {
        Builder.AddParagraph("收货单", "Title");
        Builder.AddText("单号：").AddElement("OrderNo");
        Builder.AddTable("Lines", ["MC", "Qty"], new TableFormat { CellFormat = ..., Alignment = TextAlignment.Center });
        Builder.AddImage("Logo", widthInches: 2.0, heightInches: 1.0);
    }

    // Once elements declare DataPath, MapToData / MapFromData can be omitted
    // (DataPathMapper auto-mapping in the base package, iteration 9).
}
```

### 2. Four operations

```csharp
var service = new ReceivingOrderTemplateService();

// Build the initial template (with content-control SDTs)
using var template = service.BuildInitialTemplateFile();

// Validate the template against the contract: Missing / WrongType / Ambiguous are errors,
// missing optional fields are warnings
var validation = service.Validate(templateStream);

// Pre-fill data guard: missing required fields/tables are errors,
// type mismatches / extra fields are warnings
var dataValidation = service.ValidateData(order);

// Strongly-typed fill (text / images / table rows; soft validation during fill)
using var filled = service.Fill(templateStream, order);

// Read strongly-typed data back from the filled template (including table rows)
var parsed = service.Parse(filledStream);
```

### 3. Three-step loop and key conventions

- Build → fill → read-back all share the same **tag-based location** logic (`SdtLocator`, body/header/footer); control tags must be unique document-wide.
- Tables use a "sample row" (one SDT per cell): on fill the sample row is deep-copied N times and filled row by row; **each cloned SDT gets a fresh unique `w:id`**.
- Image fill adds an image part + relationship to the package for a new `rId` and replaces `<a:blip r:embed>`; size/position/wrap are inherited from the placeholder.
- Soft validation during fill: `Drifted`/`Extra` are recorded as warnings; missing required elements follow the policy (throw by default, configurable via `SkipAndWarn`).

## Concrete builders: plugin capability = typed methods

`ITemplateBuilder` keeps only a single `Save` (the framework persists the contract). All layout capabilities are exposed as **methods on the concrete plugin builder**: declaring `TemplateService<TData, TBuilder>` is equivalent to declaring "which plugin I use"; inside the parameterless `BuildInitialTemplate()` you call the typed `Builder` instance directly — maximum freedom.

```csharp
public sealed class DeliveryOrderTemplateService : TemplateService<DeliveryOrderData, WordTemplateBuilder>
{
    protected override void BuildInitialTemplate()
    {
        Builder.SetPageSetup(new PageSetup { Size = PageSize.A5, Orientation = PageOrientation.Landscape });
        Builder.AddHeader(BuildHeader);      // AddHeader(Action<WordTemplateBuilder>)
        Builder.AddFooter(BuildFooter);
        Builder.AddTable("Lines", ["行号", "物料名称", "数量", "单位"],
            new TableFormat { HeaderFormat = ..., CellFormat = ..., Alignment = TextAlignment.Center,
                              ColumnWidthsCm = [1.8, 8.5, 3.2, 3.0] });
    }
}
```

`WordTemplateBuilder` methods: `SetPageSetup` / `AddHeader` / `AddFooter` / `AddLayoutTable` / `AddCell`
(header/footer "left-center-right" three columns) / `AddParagraph` / `AddText` / `AddElement` / `AddTable` / `AddImage` / `AddPageNumber`
(default renders "第x页，总x页"). Other plugins define their own builder classes.

## Samples

`samples/TemplateFrame.Demo.Word` is a full **delivery order** demo (`DeliveryOrderTemplateService`, A5 landscape):
- **Two-layer header**: brand layer (company LOGO | delivery order title | QR code + page number below); document header layer (order no + supplier per half line / order date + operator + remark at 1:1:2)
- **Detail table**: row no / material code / material name / unit / planned qty / actual qty / batch no / supplier batch no / warehouse (9 columns, explicit widths, narrow centered row-no column)
- **Two-line footer**: planned delivery date / actual arrival date + receiver
- **Pre-/post-receipt fills**: before receipt, actual arrival date, receiver, actual qty, batch no, warehouse are empty; filled in after receipt
- **Full loop**: build → validate → fill (pre/post receipt) → read-back — the read-back step reads the filled docx (post-receipt is the focus) → `service.Parse` → prints strongly-typed `DeliveryOrderData` (9-column multi-row detail, empty fields shown)

```bash
dotnet run --project samples/TemplateFrame.Demo.Word
```

Output goes to the system temp directory `%TEMP%\TemplateFrame.Demo.Word` by default (override with a command-line argument), producing `Word-DeliveryOrder-template` / `Word-DeliveryOrder-pre` / `Word-DeliveryOrder-post` docx files. The QR code is generated by the demo with QRCoder as PNG; the company LOGO is a demo-generated placeholder PNG — the framework replaces images, it does not generate them.

`samples/TemplateFrame.Demo.Excel` is the **delivery order Excel version** (`DeliveryOrderExcelTemplateService`, reuses the same contract and `FillData` mapping; no page setup: 3×9 grid header (LOGO left / title center / QR right) / 9-column detail / defined-name location):

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel
```

Output goes to `%TEMP%\TemplateFrame.Demo.Excel`, producing `Excel-DeliveryOrder-template` / `-pre` / `-post` xlsx files with the same build → validate → fill (pre/post receipt) → read-back loop.

`src/TemplateFrame.Excel.Simple` is the **simplified Excel plugin**: it only supports "header row + data rows" table import/export (`SimpleExcel.Write` / `SimpleExcel.Read`), marks the table position with a defined name (default `TF_Table`), and suits most list-style data. No merges / images / page setup (see [plugin README](src/TemplateFrame.Excel.Simple/README.md)).

`samples/TemplateFrame.Demo.Excel.Simple` is the **materials data demo** (`SimpleExcel` template → fill → parse back; headers: code / name / base unit / package spec / model):

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel.Simple
```

Output goes to `%TEMP%\TemplateFrame.Demo.Excel.Simple`:
- `Excel-Simple-Materials-template.xlsx`: **template** (headers only, defines the column structure)
- `Excel-Simple-Materials-filled.xlsx`: **filled** (headers + material data rows)
- Console prints the **parsed-back** result (reads the filled file → `SimpleExcel.Read` → prints headers and each row)

## Build and test

```bash
dotnet build TemplateFrame.slnx
dotnet test  TemplateFrame.slnx
```

## Packaging

```bash
dotnet pack src/TemplateFrame/TemplateFrame.csproj -c Release -o artifacts
dotnet pack src/TemplateFrame.Word/TemplateFrame.Word.csproj -c Release -o artifacts
dotnet pack src/TemplateFrame.Excel/TemplateFrame.Excel.csproj -c Release -o artifacts
dotnet pack src/TemplateFrame.Excel.Simple/TemplateFrame.Excel.Simple.csproj -c Release -o artifacts
```

Packages include XML documentation and the README; symbol packages (snupkg) are also produced. Versioning and publishing: see [docs/PUBLISHING.md](docs/PUBLISHING.md).