using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Internal;
using Sr = TemplateFrame.Excel.Simple.Localization.Sr;

namespace TemplateFrame.Excel.Simple;

internal static class SimpleExcelContractReader
{
    public static FillData Read(Stream source, TemplateContract contract, SimpleExcelOptions? options = null)
    {
        Guard.ThrowIfNull(source, nameof(source));
        Guard.ThrowIfNull(contract, nameof(contract));
        options ??= new SimpleExcelOptions();
        Guard.ThrowIfNull(options.ResourceLimits, nameof(options.ResourceLimits));
        options.ResourceLimits.Validate();
        var table = SimpleExcelContract.RequireSingleTable(contract);

        try
        {
            using var buffered = source.CanSeek ? null : new MemoryStream(StreamUtil.ReadAllBytes(source, options.ResourceLimits), writable: false);
            return ReadCore(buffered ?? source, table, options, packageChecked: buffered is not null);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or FileFormatException
                                   or XmlException or FormatException or OverflowException)
        {
            throw new InvalidOperationException(Sr.Get("SimpleExcel.Contract.CannotOpen", ex.Message), ex);
        }
    }

    private static FillData ReadCore(Stream source, TableElement table, SimpleExcelOptions options, bool packageChecked)
    {
        if (source.CanSeek)
        {
            source.Position = 0;
        }

        if (!packageChecked)
        {
            StreamUtil.ValidatePackage(source, options.ResourceLimits);
        }
        using (var document = SpreadsheetDocument.Open(source, false,
                   new OpenSettings { MaxCharactersInPart = options.ResourceLimits.MaxXmlCharactersPerPart }))
        {
            var layout = DefinedNameLayoutResolver.Resolve(document, options, table);
            if (layout is not null)
            {
                return ReadByDefinedNames(document, options, table, layout);
            }
        }

        return ReadByHeaderText(source, table, options);
    }

    private static FillData ReadByHeaderText(Stream source, TableElement table, SimpleExcelOptions options)
    {
        if (source.CanSeek)
        {
            source.Position = 0;
        }

        var loaded = SimpleExcelReader.Read(source, options.TableName, options.ResourceLimits, packageChecked: true);
        var headers = new SimpleExcelHeaderLookup(table, loaded.Headers);

        var rows = new List<IReadOnlyDictionary<string, object?>>();
        for (var r = 0; r < loaded.Rows.Count; r++)
        {
            var rowValues = new Dictionary<string, object?>();
            for (var c = 0; c < loaded.Headers.Count && c < loaded.Rows[r].Count; c++)
            {
                var header = loaded.Headers[c]?.Trim();
                if (header is { Length: > 0 } && headers.Columns.TryGetValue(header, out var column))
                {
                    rowValues[column.Key] = loaded.Rows[r][c];
                }
            }

            rows.Add(rowValues);
        }

        return new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                [table.Key] = rows,
            },
        };
    }

    private static FillData ReadByDefinedNames(
        SpreadsheetDocument document,
        SimpleExcelOptions options,
        TableElement table,
        DefinedNameLayout layout)
    {
        var workbookPart = document.WorkbookPart!;
        var worksheetPart = SimpleExcelReader.ResolveWorksheetPart(workbookPart, layout.Range.Sheet);
        var rows = worksheetPart?.Worksheet?.GetFirstChild<SheetData>()?.Elements<Row>().ToList() ?? [];
        var sharedStrings = SimpleExcelReader.MaterializeSharedStrings(workbookPart);
        var rowLookup = SimpleExcelReader.BuildRowLookup(rows);

        var budget = new TemplateResourceBudget(options.ResourceLimits);
        budget.AddCells(layout.Headers.Count);
        var dataRows = new List<IReadOnlyDictionary<string, object?>>();
        var cells = new SimpleExcelRowCells();
        var dateStyles = new SimpleExcelDateStyles(workbookPart);
        foreach (var r in rowLookup.Keys.Where(r => r > layout.Range.StartRow).OrderBy(r => r))
        {
            budget.AddRows(1, table.Columns.Count);
            cells.Load(rowLookup[r]);
            var rowValues = new Dictionary<string, object?>();
            var any = false;
            foreach (var column in table.Columns)
            {
                if (layout.ColumnIndex.TryGetValue(column.Key, out var col))
                {
                    var cell = cells.Find(r, col);
                    var value = cell is null ? null : SimpleExcelReader.ReadCellValue(dateStyles, sharedStrings, cell);
                    rowValues[column.Key] = value;
                    any |= value is not null;
                }
                else
                {
                    rowValues[column.Key] = null;
                }
            }

            if (any)
            {
                dataRows.Add(rowValues);
            }
        }

        return new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                [table.Key] = dataRows,
            },
        };
    }

}
