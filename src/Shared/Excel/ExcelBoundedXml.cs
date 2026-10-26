using System.Text;
using System.Xml;

namespace TemplateFrame.Internal.Excel;

internal static class ExcelBoundedXml
{
    internal static string ReadElement(XmlReader reader, int maximum, CancellationToken cancellation)
    {
        using var output = new BoundedTextWriter(maximum, cancellation);
        using (var subtree = reader.ReadSubtree())
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { OmitXmlDeclaration = true }))
        {
            writer.WriteNode(subtree, false);
        }
        return output.ToString();
    }

    private sealed class BoundedTextWriter(int maximum, CancellationToken cancellation) : TextWriter
    {
        private readonly StringBuilder _text = new();
        public override Encoding Encoding => Encoding.Unicode;

        public override void Write(char value)
        {
            Check(1);
            _text.Append(value);
        }

        public override void Write(char[] buffer, int index, int count)
        {
            Check(count);
            _text.Append(buffer, index, count);
        }

        public override void Write(string? value)
        {
            Check(value?.Length ?? 0);
            _text.Append(value);
        }

        private void Check(int count)
        {
            cancellation.ThrowIfCancellationRequested();
            if (count > maximum - _text.Length)
                throw new InvalidOperationException("Streaming import row or shared string item exceeds MaxRowCharacters.");
        }

        public override string ToString() => _text.ToString();
    }
}
