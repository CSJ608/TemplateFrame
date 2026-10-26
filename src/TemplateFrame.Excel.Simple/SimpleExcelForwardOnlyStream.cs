namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelForwardOnlyStream(Stream target) : Stream
{
    private long _position;
    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => target.CanWrite;
    public override long Length => _position;
    public override long Position { get => _position; set => throw new NotSupportedException(); }
    public override void Flush() => target.Flush();
    public override void Write(byte[] buffer, int offset, int count)
    {
        target.Write(buffer, offset, count);
        _position += count;
    }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
}
