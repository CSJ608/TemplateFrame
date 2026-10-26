using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelRowCells
{
    private readonly Dictionary<string, Cell> _cells = new(StringComparer.Ordinal);
    private OpenXmlElement? _next;

    internal void Load(Row row)
    {
        _cells.Clear();
        _next = row.FirstChild;
    }

    internal Cell? Find(int rowIndex, int columnIndex)
    {
        var reference = SimpleExcelAddress.CellReference((uint)rowIndex, columnIndex);
        if (_cells.TryGetValue(reference, out var found))
        {
            return found;
        }

        while (_next is { } element)
        {
            _next = element.NextSibling();
            if (element is not Cell cell || cell.CellReference?.Value is not { } cellReference)
            {
                continue;
            }

            if (!_cells.ContainsKey(cellReference)) _cells.Add(cellReference, cell);
            if (cellReference == reference) return cell;
        }

        return null;
    }
}
