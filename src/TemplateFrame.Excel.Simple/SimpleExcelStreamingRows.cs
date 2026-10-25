using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelStreamingRows : IDisposable
{
    private readonly IEnumerator<IReadOnlyList<object?>> _source;
    private readonly SimpleExcelStreamingLayout _layout;
    private readonly CancellationToken _cancellation;
    private bool _completed;
    internal int LastRow { get; private set; }

    internal SimpleExcelStreamingRows(IEnumerable<IReadOnlyList<object?>> source, SimpleExcelStreamingLayout layout, CancellationToken cancellation)
    {
        cancellation.ThrowIfCancellationRequested();
        _layout = layout;
        _cancellation = cancellation;
        LastRow = layout.StartRow;
        _source = source.GetEnumerator();
    }

    internal Row Header()
        => new(_layout.Headers.Select((header, index) => new Cell
        {
            CellReference = SimpleExcelAddress.CellReference((uint)_layout.StartRow, _layout.StartColumn + index),
            StyleIndex = _layout.Options.BoldHeader ? 1u : 0u,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(header) { Space = SpaceProcessingModeValues.Preserve }),
        }))
        { RowIndex = (uint)_layout.StartRow };

    internal bool TryRead(out Row row)
    {
        row = new Row { RowIndex = (uint)(LastRow + 1) };
        _cancellation.ThrowIfCancellationRequested();
        if (_completed) return false;
        if (!_source.MoveNext())
        {
            _completed = true;
            return false;
        }
        _cancellation.ThrowIfCancellationRequested();
        _layout.Budget.AddRows(1, _layout.Headers.Length);
        var values = _source.Current ?? throw new ArgumentException("A source row cannot be null.");
        for (var column = 0; column < _layout.Headers.Length; column++)
        {
            _cancellation.ThrowIfCancellationRequested();
            var value = column < values.Count ? values[column] : null;
            if (value is null) continue;
            if (LastRow == SimpleExcelAddress.MaxRow)
                throw new InvalidOperationException("Rows exceed the worksheet boundary.");
            var cell = new Cell { CellReference = SimpleExcelAddress.CellReference((uint)(LastRow + 1), _layout.StartColumn + column) };
            SimpleExcelCellWriter.Write(cell, value, 2);
            row.Append(cell);
        }
        if (row.HasChildren) LastRow++;
        return true;
    }

    public void Dispose() => _source.Dispose();
}
