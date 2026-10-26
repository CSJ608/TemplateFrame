using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using TemplateFrame.Internal.Excel;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class BoundedXmlReuseTests
{
    private const string Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static IEnumerable<object[]> Samples()
    {
        yield return ["<si><t>中文😀 &amp; &lt; &gt; &#13;\n</t></si><si/><si><t/></si>"];
        yield return ["<si><r><rPr><b/></rPr><t>rich</t></r><r><t>text</t></r></si><si><t>direct</t><r><t>ignored</t></r></si>"];
        yield return ["<si xml:space='preserve' xml:lang='en'><t> a </t><!--comment--><?pi value?></si><si><t><![CDATA[a<b]]></t></si>"];
        yield return ["<s:si xmlns:s='" + Namespace + "'><s:t>prefix</s:t></s:si><si xmlns:q='urn:a' q:flag='x'><t>one</t></si><si xmlns:q='urn:b' q:flag='y'><t>two</t></si>"];
        yield return ["<si><x:extra xmlns:x='urn:a'><x:inner xmlns:x='urn:b'/></x:extra><t>text</t></si><si><t>next</t></si>"];
        yield return ["<si><t>" + new string('中', 32767) + "</t></si><si><t>x</t></si>"];
        yield return ["<si></si><si><t></t></si><si><t xml:space='preserve'> \t </t></si>"];
    }

    [Theory]
    [MemberData(nameof(Samples))]
    public void ReuseMatchesFreshSerializationAndBothStringSemantics(string items)
    {
        var expected = Read(items, reuse: false);
        var actual = Read(items, reuse: true);
        Assert.Equal(expected, actual);
        for (var index = 0; index < expected.Count; index++)
        {
            var first = new SharedStringItem(expected[index]);
            var second = new SharedStringItem(actual[index]);
            Assert.Equal(first.Text?.Text, second.Text?.Text);
            Assert.Equal(SimpleExcelReader.GetSharedStringText(first), SimpleExcelReader.GetSharedStringText(second));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BudgetAppliesToEachSerializedElementIncludingNamespaceAndEscaping(bool reuse)
    {
        const string item = "<si><t>&amp;&lt;&gt;中文😀</t></si>";
        var expected = Read(item, false)[0];
        Assert.Equal(new[] { expected, expected }, Read(item + item, reuse, expected.Length));
        Assert.Throws<InvalidOperationException>(() => Read(item, reuse, expected.Length - 1));
    }

    [Fact]
    public void InheritedNamespacesAndXmlSpaceMatchFreshSerialization()
    {
        var document = "<sst xmlns='" + Namespace + "' xmlns:s='" + Namespace
            + "' xmlns:p='urn:parent' xml:space='preserve'><s:si p:flag='one'><s:t> a </s:t></s:si>"
            + "<si p:flag='two'><t> b </t></si></sst>";
        Assert.Equal(ReadDocument(document, false, 1000), ReadDocument(document, true, 1000));
    }

    [Fact]
    public void ManyNamespaceChangesDoNotLeakBetweenItems()
    {
        var items = string.Concat(Enumerable.Range(0, 2000).Select(index =>
            "<si xmlns:p='urn:item:" + index + "' p:flag='value'><p:extra/><t>" + index + "</t></si>"));
        Assert.Equal(Read(items, false), Read(items, true));
    }

    [Fact]
    public void CancellationBetweenItemsDoesNotConsumeNextElement()
    {
        using var cancellation = new CancellationTokenSource();
        using var reader = CreateReader("<si><t>first</t></si><si><t>second</t></si>");
        using var elements = new ExcelBoundedXml.ElementReader(1000, cancellation.Token);
        Assert.True(reader.ReadToFollowing("si", Namespace));
        _ = elements.ReadElement(reader);
        Assert.True(reader.ReadToFollowing("si", Namespace));
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => elements.ReadElement(reader));
        Assert.Equal(XmlNodeType.Element, reader.NodeType);
        Assert.Equal("si", reader.LocalName);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CorruptLaterElementIsNotHiddenByReuse(bool reuse)
        => Assert.Throws<XmlException>(() => Read("<si><t>valid</t></si><si><t>broken</si>", reuse));

    private static List<string> Read(string items, bool reuse, int maximum = 1024 * 1024)
        => ReadDocument("<sst xmlns='" + Namespace + "'>" + items + "</sst>", reuse, maximum);

    private static List<string> ReadDocument(string document, bool reuse, int maximum)
    {
        using var reader = XmlReader.Create(new StringReader(document),
            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        using var elements = reuse ? new ExcelBoundedXml.ElementReader(maximum, default) : null;
        var values = new List<string>();
        while (reader.Read())
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si" && reader.Depth == 1)
                values.Add(elements is null ? ExcelBoundedXml.ReadElement(reader, maximum, default) : elements.ReadElement(reader));
        return values;
    }

    private static XmlReader CreateReader(string items)
        => XmlReader.Create(new StringReader("<sst xmlns='" + Namespace + "'>" + items + "</sst>"),
            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
}
