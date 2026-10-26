using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using TemplateFrame.Internal.Excel;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class SharedStringReparseTests
{
    private const string Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static IEnumerable<object[]> Items()
    {
        yield return ["<si/>"];
        yield return ["<si><t/></si>"];
        yield return ["<si><t></t></si>"];
        yield return ["<si><t>first</t><t>second</t></si>"];
        yield return ["<si><r><t>one</t><t>ignored</t></r><r><t>two</t></r></si>"];
        yield return ["<si><r><t>ignored</t></r><t>direct</t></si>"];
        yield return ["<si><t/><r><t>ignored</t></r></si>"];
        yield return ["<si><r/><rPh sb='0' eb='1'><t>phonetic</t></rPh></si>"];
        yield return ["<si><t xml:space='preserve'> \t a \r\n </t></si>"];
        yield return ["<si><t>&amp;&lt;&gt;&quot;&apos;&#10;中文😀</t></si>"];
        yield return ["<si><t><![CDATA[a<b]]>tail<!--comment--><?pi value?></t></si>"];
        yield return ["<si><x:t>foreign</x:t><t>direct</t></si>"];
        yield return ["<si><x:extra><t>nested</t></x:extra><r><t>run</t></r></si>"];
        yield return ["<si><t xmlns=''>no namespace</t></si>"];
        yield return ["<s:si><s:t>prefixed</s:t></s:si>"];
        yield return ["<si mc:Ignorable='x'><x:extra/><t>known</t></si>"];
        yield return ["<si><mc:AlternateContent><mc:Choice Requires='x'><t>choice</t></mc:Choice><mc:Fallback><t>fallback</t></mc:Fallback></mc:AlternateContent></si>"];
        yield return ["<si><t><x:extra/>tail</t></si>"];
        yield return ["<si><t>" + new string('中', 32767) + "</t></si>"];
    }

    [Theory]
    [MemberData(nameof(Items))]
    public void BoundedReparseMatchesWholeTableSdkSelectors(string item)
    {
        var xml = "<sst xmlns='" + Namespace + "' xmlns:s='" + Namespace
            + "' xmlns:x='urn:extension' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xml:space='preserve'>"
            + item + "<si><t>following</t></si></sst>";
        var expected = new SharedStringTable(xml).Elements<SharedStringItem>().ToArray();
        using var reader = XmlReader.Create(new StringReader(xml),
            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        using var elements = new ExcelBoundedXml.ElementReader(1024 * 1024, default);
        var index = 0;
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1
                || reader.LocalName != "si" || reader.NamespaceURI != Namespace) continue;
            var actual = new SharedStringItem(elements.ReadElement(reader));
            Assert.Equal(expected[index].Text?.Text, actual.Text?.Text);
            Assert.Equal(SimpleExcelReader.GetSharedStringText(expected[index]), SimpleExcelReader.GetSharedStringText(actual));
            index++;
        }
        Assert.Equal(expected.Length, index);
    }

    [Theory]
    [InlineData("<si/>", null, "")]
    [InlineData("<si><t/></si>", "", "")]
    [InlineData("<si><t>first</t><t>second</t></si>", "first", "first")]
    [InlineData("<si><r><t>one</t><t>ignored</t></r><r><t>two</t></r></si>", null, "onetwo")]
    [InlineData("<si><r><t>ignored</t></r><t>direct</t></si>", "direct", "direct")]
    public void SelectorContractDistinguishesFirstDirectTextFromRuns(string xml, string? direct, string simple)
    {
        var item = new SharedStringItem(xml.Replace("<si", "<si xmlns='" + Namespace + "'"));
        Assert.Equal(direct, item.Text?.Text);
        Assert.Equal(simple, SimpleExcelReader.GetSharedStringText(item));
    }

    [Fact]
    public void ExistingSerializationNormalizesCarriageReturnEntityUnlikeWholeTable()
    {
        var xml = "<sst xmlns='" + Namespace + "'><si><t>a&#13;&#10;b</t></si></sst>";
        Assert.Equal("a\r\nb", new SharedStringTable(xml).Elements<SharedStringItem>().Single().Text!.Text);
        foreach (var reuse in new[] { false, true })
        {
            using var reader = XmlReader.Create(new StringReader(xml));
            Assert.True(reader.ReadToFollowing("si", Namespace));
            using var elements = new ExcelBoundedXml.ElementReader(1000, default);
            var serialized = reuse ? elements.ReadElement(reader) : ExcelBoundedXml.ReadElement(reader, 1000, default);
            var item = new SharedStringItem(serialized);
            Assert.Equal("a\nb", item.Text!.Text);
            Assert.Equal("a\nb", SimpleExcelReader.GetSharedStringText(item));
        }
    }
}
