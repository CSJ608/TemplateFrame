using System.Diagnostics;
using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Excel;
using TemplateFrame.Excel.Simple;
using TemplateFrame.Word;

namespace TemplateFrame.CapacityProbe;

internal sealed class Workload
{
    private readonly ProbeSettings _settings;
    private readonly TemplateContract _contract;
    private readonly string[] _keys;
    private readonly string _outputPath;
    private readonly SimpleExcelTable? _table;
    private readonly FillData? _data;
    private readonly byte[]? _template;
    private object? _result;
    private readonly string _scratchDirectory;
    private int _streamedRows;

    public long ScratchPeakBytes { get; private set; }
    public int ScratchRemainingFiles => Directory.Exists(_scratchDirectory) ? Directory.GetFiles(_scratchDirectory).Length : 0;

    public long OutputBytes => File.Exists(_outputPath) ? new FileInfo(_outputPath).Length : 0;

    public Workload(ProbeSettings settings, int worker)
    {
        _settings = settings;
        _scratchDirectory = Path.Combine(settings.Directory, $"scratch-{worker}");
        _keys = Enumerable.Range(0, settings.Columns).Select(c => "C" + c).ToArray();
        _outputPath = Path.Combine(settings.Directory, $"output-{worker}.{(settings.Mode == "word-write" ? "docx" : "xlsx")}");
        var columns = _keys.Select((key, c) => new TextElement
        {
            Key = key,
            ValueType = settings.Mode == "word-write" ? typeof(string) : (c % 6) switch
            {
                3 => typeof(decimal),
                5 => typeof(DateTime),
                _ => typeof(string),
            },
            Format = c % 6 == 5 ? "yyyy-MM-dd" : null,
        }).ToArray();
        _contract = new TemplateContract { Elements = [new TableElement { Key = "Items", Columns = columns }] };
        if (settings.Mode == "template-streaming-write")
        {
            using var streamingTemplate = new MemoryStream();
            var builder = new ExcelTemplateBuilder();
            builder.AddTable("Items", _keys, startCell: "A1");
            builder.Save(streamingTemplate);
            _template = streamingTemplate.ToArray();
            return;
        }
        if (!settings.Mode.EndsWith("write", StringComparison.Ordinal) || settings.Mode.EndsWith("streaming-write", StringComparison.Ordinal)) return;
        var rows = Enumerable.Range(0, settings.Rows).Select(r => (IReadOnlyList<object?>)
            Enumerable.Range(0, settings.Columns).Select(c => Value(r, c, settings.Mode == "word-write")).ToArray()).ToArray();
        if (settings.Mode == "simple-write")
        {
            _table = new SimpleExcelTable { Headers = _keys, Rows = rows };
            return;
        }
        _data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Items"] = rows.Select(row => (IReadOnlyDictionary<string, object?>)
                    _keys.Select((key, c) => (key, Value: row[c])).ToDictionary(pair => pair.key, pair => pair.Value)).ToArray(),
            },
        };
        using var template = new MemoryStream();
        if (settings.Mode == "word-write")
        {
            _contract = _contract with { Elements = [.. _contract.Elements, new ImageElement { Key = "Logo" }, new ImageElement { Key = "Stamp" }] };
            _data = new FillData { Tables = _data.Tables, Values = new Dictionary<string, object?> { ["Logo"] = ProbeImage.Bytes, ["Stamp"] = ProbeImage.Bytes } };
            var builder = new WordTemplateBuilder();
            builder.AddParagraph("合成容量验证单据");
            builder.AddImage("Logo", widthInches: 0.8, heightInches: 0.8);
            builder.AddTable("Items", _keys);
            builder.AddImage("Stamp", widthInches: 0.8, heightInches: 0.8);
            builder.Save(template);
        }
        else
        {
            var builder = new ExcelTemplateBuilder();
            builder.AddTable("Items", _keys, startCell: "A1");
            builder.Save(template);
        }
        _template = template.ToArray();
    }

    public double Run()
    {
        var timer = Stopwatch.StartNew();
        if (_settings.Mode.EndsWith("read", StringComparison.Ordinal))
        {
            using var input = File.OpenRead(_settings.InputPath);
            _result = _settings.Mode switch
            {
                "simple-read" => SimpleExcel.Read(input, "Items", _settings.Limits),
                "streaming-read" => ReadStreaming(input),
                "contract-streaming-read" => ReadContractStreaming(input),
                "template-streaming-read" => ReadTemplateStreaming(input),
                "contract-read" => SimpleExcelContract.Read(input, _contract, new SimpleExcelOptions { TableName = "Items", ResourceLimits = _settings.Limits }),
                _ => new ExcelTemplateParser(_settings.Limits, null).Parse(input, _contract),
            };
        }
        else
        {
            using var target = File.Create(_outputPath);
            if (_settings.Mode == "simple-write")
                SimpleExcel.Write(target, _table!, new SimpleExcelOptions { TableName = "Items", ResourceLimits = _settings.Limits }, _keys);
            else if (_settings.Mode == "template-streaming-write")
            {
                using var template = new MemoryStream(_template!, writable: false);
                new ExcelTemplateFiller(new TemplateFillOptions { ResourceLimits = _settings.Limits }).FillStreaming(template, target, _contract,
                    new ExcelStreamingData
                    {
                        Tables = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, object?>>> { ["Items"] = StreamTemplateRows() },
                    });
            }
            else if (_settings.Mode.EndsWith("streaming-write", StringComparison.Ordinal))
                SimpleExcel.WriteStreaming(target, new SimpleExcelStreamingTable { Headers = _keys, ColumnKeys = _keys, Rows = StreamRows() }, new SimpleExcelStreamingOptions
                {
                    WorkbookOptions = new SimpleExcelOptions { TableName = "Items", ResourceLimits = _settings.Limits },
                    ColumnWidths = _settings.Mode == "fixed-streaming-write" ? Enumerable.Repeat(20d, _settings.Columns).ToArray() : null,
                });
            else
            {
                using var template = new MemoryStream(_template!, writable: false);
                var options = new TemplateFillOptions { ResourceLimits = _settings.Limits };
                using var output = (_settings.Mode == "word-write"
                    ? new WordTemplateFiller(options).Fill(template, _contract, _data!)
                    : new ExcelTemplateFiller(options).Fill(template, _contract, _data!)).Output;
                output.CopyTo(target);
            }
        }
        return timer.Elapsed.TotalMilliseconds;
    }

    public void Validate()
    {
        if (_result is ExcelStreamingParseResult parsed)
        {
            if (!parsed.Completed || parsed.RowsRead != _settings.Rows || parsed.ValuesRead != 0
                || _streamedRows != _settings.Rows || ScratchRemainingFiles != 0)
                throw new InvalidDataException("Template streaming read shape or cleanup mismatch.");
        }
        else if (_result is SimpleExcelStreamingReadResult streamed)
        {
            if (!streamed.Completed || streamed.RowsRead != _settings.Rows || _streamedRows != _settings.Rows
                || !streamed.Headers.SequenceEqual(_keys) || ScratchRemainingFiles != 0)
                throw new InvalidDataException("Streaming read shape or cleanup mismatch.");
        }
        else if (_result is SimpleExcelTable table)
        {
            if (table.Rows.Count != _settings.Rows || table.Headers.Count != _settings.Columns) throw new InvalidDataException("Read shape mismatch.");
            for (var r = 0; r < table.Rows.Count; r++)
                for (var c = 0; c < _settings.Columns; c++) Check(table.Rows[r][c], r, c);
        }
        else if (_result is FillData data)
        {
            var rows = data.Tables["Items"];
            if (rows.Count != _settings.Rows) throw new InvalidDataException("Contract row count mismatch.");
            for (var r = 0; r < rows.Count; r++)
                for (var c = 0; c < _settings.Columns; c++) Check(rows[r][_keys[c]], r, c);
        }
        else OutputValidation.Check(_outputPath, _settings);
        GC.KeepAlive(_result);
    }

    private SimpleExcelStreamingReadResult ReadStreaming(Stream input)
    {
        Directory.CreateDirectory(_scratchDirectory);
        var result = SimpleExcel.ReadStreaming(input, row =>
        {
            if (row.RowNumber != _streamedRows + 2 || row.Values.Count != _settings.Columns)
                throw new InvalidDataException("Streaming row shape mismatch.");
            for (var column = 0; column < _settings.Columns; column++) Check(row.Values[column], _streamedRows, column);
            _streamedRows++;
            return true;
        }, new SimpleExcelStreamingReadOptions
        {
            TableName = "Items",
            ResourceLimits = _settings.Limits,
            TemporaryDirectory = _scratchDirectory,
            MaxTemporaryBytes = 1024L * 1024 * 1024,
        });
        ScratchPeakBytes = result.TemporaryBytesWritten;
        return result;
    }

    private SimpleExcelStreamingReadResult ReadContractStreaming(Stream input)
    {
        Directory.CreateDirectory(_scratchDirectory);
        var result = SimpleExcelContract.ReadStreaming(input, _contract, row =>
        {
            if (row.RowNumber != _streamedRows + 2 || row.Values.Count != _settings.Columns)
                throw new InvalidDataException("Contract streaming row shape mismatch.");
            for (var column = 0; column < _settings.Columns; column++) Check(row.Values[_keys[column]], _streamedRows, column);
            _streamedRows++;
            return true;
        }, new SimpleExcelStreamingReadOptions
        {
            TableName = "Items",
            ResourceLimits = _settings.Limits,
            TemporaryDirectory = _scratchDirectory,
            MaxTemporaryBytes = 1024L * 1024 * 1024,
        });
        ScratchPeakBytes = result.TemporaryBytesWritten;
        return result;
    }

    private ExcelStreamingParseResult ReadTemplateStreaming(Stream input)
    {
        Directory.CreateDirectory(_scratchDirectory);
        var result = new ExcelTemplateParser().ParseStreaming(input, _contract, item =>
        {
            if (item is not ExcelStreamingParseRow row || row.Key != "Items"
                || row.DataRowNumber != _streamedRows + 1 || row.Values.Count != _settings.Columns)
                throw new InvalidDataException("Template streaming row shape mismatch.");
            for (var column = 0; column < _settings.Columns; column++)
            {
                var actual = row.Values[_keys[column]];
                var expected = Value(_streamedRows, column, false);
                if (!Equals(actual, expected)) throw new InvalidDataException($"Template typed value mismatch at {_streamedRows}:{column}.");
            }
            _streamedRows++;
            return true;
        }, new ExcelStreamingParseOptions
        {
            ResourceLimits = _settings.Limits,
            TemporaryDirectory = _scratchDirectory,
            MaxTemporaryBytes = 1024L * 1024 * 1024,
        }, warning: issue => throw new InvalidDataException("Unexpected template warning: " + issue.Message));
        ScratchPeakBytes = result.TemporaryBytesWritten;
        return result;
    }

    private static void Check(object? actual, int row, int column)
    {
        var expected = Value(row, column, false);
        var valid = expected is decimal number ? Convert.ToDecimal(actual, CultureInfo.InvariantCulture) == number : Equals(actual, expected);
        if (!valid) throw new InvalidDataException($"Value mismatch at {row}:{column}.");
    }

    private IEnumerable<IReadOnlyList<object?>> StreamRows()
    {
        for (var row = 0; row < _settings.Rows; row++)
            yield return Enumerable.Range(0, _settings.Columns).Select(column => Value(row, column, false)).ToArray();
    }

    private IEnumerable<IReadOnlyDictionary<string, object?>> StreamTemplateRows()
    {
        for (var row = 0; row < _settings.Rows; row++)
            yield return _keys.Select((key, column) => (key, Value: Value(row, column, false)))
                .ToDictionary(pair => pair.key, pair => (object?)pair.Value);
    }

    internal static object Value(int row, int column, bool textOnly) => !textOnly && column % 6 == 3 ? 10m + row
        : !textOnly && column % 6 == 5 ? new DateTime(2026, 1, 1).AddDays(row % 365)
        : $"物料{row:D7}-字段{column:D2}";

}
