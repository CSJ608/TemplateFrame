namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelTemporaryStorage(SimpleExcelStreamingReadOptions options) : IDisposable
{
    private readonly List<FileStream> _files = new();
    private long _reserved;

    internal long ReservedBytes => _reserved;

    internal FileStream CreateFile()
    {
        var root = options.TemporaryDirectory ?? Path.GetTempPath();
        var path = Path.Combine(root, "TemplateFrame-" + Guid.NewGuid().ToString("N") + ".tmp");
        var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 8192, FileOptions.DeleteOnClose);
        _files.Add(file);
        return file;
    }

    internal void Reserve(long bytes)
    {
        if (bytes < 0) throw new ArgumentOutOfRangeException(nameof(bytes));
        if (bytes > options.MaxTemporaryBytes - _reserved)
            throw new InvalidOperationException("Streaming import exceeds MaxTemporaryBytes.");
        _reserved += bytes;
    }

    internal Stream BufferInput(Stream source, CancellationToken cancellation)
    {
        var file = CreateFile();
        var buffer = new byte[81920];
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            var count = source.Read(buffer, 0, buffer.Length);
            cancellation.ThrowIfCancellationRequested();
            if (count == 0) break;
            if (count > options.ResourceLimits.MaxInputBytes - file.Length)
                throw new InvalidOperationException("Streaming import exceeds MaxInputBytes.");
            Reserve(count);
            file.Write(buffer, 0, count);
        }
        file.Position = 0;
        return file;
    }

    public void Dispose()
    {
        Exception? error = null;
        foreach (var file in _files)
        {
            try { file.Dispose(); }
            catch (Exception exception) { error ??= exception; }
        }
        _files.Clear();
        if (error is not null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
    }
}
