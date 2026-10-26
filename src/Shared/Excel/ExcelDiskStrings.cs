using System.Collections;
using System.Text;

namespace TemplateFrame.Internal.Excel;

internal sealed class ExcelDiskStrings : IReadOnlyList<string>
{
    private readonly ExcelTemporaryStorage _storage;
    private readonly IExcelStreamingReadSettings _options;
    private readonly FileStream _index;
    private readonly FileStream _text;
    private readonly BinaryReader _indexReader;
    private readonly BinaryWriter _indexWriter;
    private readonly Dictionary<int, LinkedListNode<(int Index, string Text)>> _cache = new();
    private readonly LinkedList<(int Index, string Text)> _recent = new();
    private int _cacheCharacters;

    internal ExcelDiskStrings(ExcelTemporaryStorage storage, IExcelStreamingReadSettings options)
    {
        _storage = storage;
        _options = options;
        _index = storage.CreateFile();
        _text = storage.CreateFile();
        _indexReader = new BinaryReader(_index, Encoding.UTF8, leaveOpen: true);
        _indexWriter = new BinaryWriter(_index, Encoding.UTF8, leaveOpen: true);
    }

    public int Count { get; private set; }
    internal int CachedCharacters => _cacheCharacters;
    internal int CachedEntries => _cache.Count;

    internal void Add(string text)
        => AddNullable(text);

    internal void AddNullable(string? text)
    {
        if (text?.Length > 32767) throw new InvalidOperationException("Shared string exceeds the Excel cell text limit.");
        if (Count == int.MaxValue) throw new InvalidOperationException("Shared string index exceeds supported range.");
        var bytes = text is null ? Array.Empty<byte>() : Encoding.UTF8.GetBytes(text);
        _storage.Reserve(12L + bytes.Length);
        _index.Position = _index.Length;
        _text.Position = _text.Length;
        _indexWriter.Write(_text.Position);
        _indexWriter.Write(text is null ? -1 : bytes.Length);
        _text.Write(bytes, 0, bytes.Length);
        Count++;
    }

    public string this[int index] => ReadNullable(index) ?? string.Empty;

    internal string? ReadNullable(int index)
    {
        if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
        if (_cache.TryGetValue(index, out var cached))
        {
            _recent.Remove(cached);
            _recent.AddLast(cached);
            return cached.Value.Text;
        }
        _index.Position = index * 12L;
        _text.Position = _indexReader.ReadInt64();
        var length = _indexReader.ReadInt32();
        if (length == -1) return null;
        var bytes = new byte[length];
        var read = 0;
        while (read < length)
        {
            var count = _text.Read(bytes, read, length - read);
            if (count == 0) throw new EndOfStreamException();
            read += count;
        }
        var text = Encoding.UTF8.GetString(bytes);
        Cache(index, text);
        return text;
    }

    private void Cache(int index, string text)
    {
        if (_options.SharedStringCacheEntries == 0 || text.Length > _options.SharedStringCacheCharacters) return;
        while (_recent.First is { } oldest && (_cache.Count >= _options.SharedStringCacheEntries
            || text.Length > _options.SharedStringCacheCharacters - _cacheCharacters))
        {
            _cache.Remove(oldest.Value.Index);
            _cacheCharacters -= oldest.Value.Text.Length;
            _recent.RemoveFirst();
        }
        _cache.Add(index, _recent.AddLast((index, text)));
        _cacheCharacters += text.Length;
    }

    public IEnumerator<string> GetEnumerator()
    {
        for (var index = 0; index < Count; index++) yield return this[index];
    }
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
