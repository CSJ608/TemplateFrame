using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel.Simple;
using TemplateFrame.Localization;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

/// <summary>
/// 迭代 14：SimpleExcel 契约路径"每列定义名定位"——框架产物回读与表头语言解耦（语言无关），
/// 表头文本匹配作分级回退；重复定义名 Validate 报 Ambiguous。
/// </summary>
public sealed class DefinedNameColumnTests
{
    private static readonly CultureInfo En = CultureInfo.GetCultureInfo("en");

    private static TemplateContract MaterialsContract()
        => new()
        {
            Name = "Materials",
            Version = "1.0",
            Elements =
            [
                new TableElement
                {
                    Key = "Materials",
                    DisplayName = "物料清单",
                    DataPath = "Items",
                    Columns =
                    [
                        new TextElement { Key = "编码", DisplayName = "编码", DataPath = "Code", Required = true },
                        new TextElement { Key = "名称", DisplayName = "名称", DataPath = "Name", Required = true },
                        new TextElement { Key = "数量", DisplayName = "数量", DataPath = "Qty", ValueType = typeof(decimal) },
                    ],
                },
            ],
        };

    private static FillData SampleFillData()
        => new()
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Materials"] =
                [
                    new Dictionary<string, object?> { ["编码"] = "AL-6063", ["名称"] = "铝型材 6063-T5", ["数量"] = 120.5 },
                    new Dictionary<string, object?> { ["编码"] = "SS-M8", ["名称"] = "不锈钢螺栓 M8×30", ["数量"] = 500.0 },
                ],
            },
        };

    private static DefaultTemplateLocalizer EnLocalizer()
        => new(new Dictionary<string, string>
        {
            ["en:编码"] = "Code",
            ["en:名称"] = "Name",
            ["en:数量"] = "Qty",
        });

    [Fact]
    public void Write_EnglishCulture_WritesLocalizedHeaders_AndColumnDefinedNames()
    {
        using var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, SampleFillData(), MaterialsContract(), new SimpleExcelOptions { SheetName = "物料清单" }, En, EnLocalizer());
        stream.Position = 0;

        var raw = SimpleExcel.Read(stream);
        Assert.Equal(["Code", "Name", "Qty"], raw.Headers);

        using var document = SpreadsheetDocument.Open(stream, false);
        var names = document.WorkbookPart!.Workbook.DefinedNames!
            .Elements<DefinedName>()
            .Select(d => d.Name!.Value!)
            .ToList();
        Assert.Contains("TF_Table", names);
        Assert.Contains("TF_Table_编码", names);
        Assert.Contains("TF_Table_名称", names);
        Assert.Contains("TF_Table_数量", names);
    }

    [Fact]
    public void Read_EnglishWrittenFile_ByDefinedNames_ReturnsValues_WithoutLanguage()
    {
        using var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, SampleFillData(), MaterialsContract(), new SimpleExcelOptions { SheetName = "物料清单" }, En, EnLocalizer());
        stream.Position = 0;

        // 无语言读：表头是英文，文本匹配必然失败——必须靠每列定义名定位（语言无关）
        var loaded = SimpleExcelContract.Read(stream, MaterialsContract());

        var rows = loaded.Tables["Materials"];
        Assert.Equal(2, rows.Count);
        Assert.Equal("AL-6063", rows[0]["编码"]);
        Assert.Equal("铝型材 6063-T5", rows[0]["名称"]);
        Assert.Equal(120.5, Assert.IsType<double>(rows[0]["数量"]), 3);
        Assert.Equal("SS-M8", rows[1]["编码"]);
        Assert.Equal(500.0, Assert.IsType<double>(rows[1]["数量"]), 3);
    }

    [Fact]
    public void Read_FallsBackToTextMatching_WhenColumnNamesStripped()
    {
        using var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, SampleFillData(), MaterialsContract(), new SimpleExcelOptions { SheetName = "物料清单" });
        stream.Position = 0;

        // 模拟手工修改丢了每列定义名（保留 TF_Table 区域）→ 回退中文表头文本匹配
        using var stripped = WithModifiedDefinedNames(stream, names =>
        {
            var keep = names.Elements<DefinedName>().Where(d => d.Name?.Value == "TF_Table").ToList();
            names.RemoveAllChildren<DefinedName>();
            foreach (var dn in keep)
            {
                names.Append(dn);
            }
        });

        var loaded = SimpleExcelContract.Read(stripped, MaterialsContract());

        var rows = loaded.Tables["Materials"];
        Assert.Equal(2, rows.Count);
        Assert.Equal("AL-6063", rows[0]["编码"]);
        Assert.Equal("铝型材 6063-T5", rows[0]["名称"]);
        Assert.Equal(120.5, Assert.IsType<double>(rows[0]["数量"]), 3);
    }

    [Fact]
    public void Validate_ReportsAmbiguous_WhenColumnDefinedNameDuplicated()
    {
        using var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, SampleFillData(), MaterialsContract(), new SimpleExcelOptions { SheetName = "物料清单" });
        stream.Position = 0;

        using var duplicated = WithModifiedDefinedNames(stream, names =>
        {
            var target = names.Elements<DefinedName>().First(d => d.Name?.Value == "TF_Table_编码");
            names.Append(new DefinedName { Name = target.Name?.Value, Text = target.Text });
        });

        var result = SimpleExcelContract.Validate(duplicated, MaterialsContract());

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Ambiguous && i.Key == "编码");
        Assert.DoesNotContain(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing && i.Key == "编码");
        Assert.DoesNotContain(result.Issues, i => i.Code == TemplateValidationIssueCode.Extra && i.Key == "编码");
    }

    /// <summary>
    /// 2.1.1 回归：两个不同列的定义名指向同一单元格（各自只出现一次，非重名）——
    /// Validate 应报 Ambiguous（两列、位置歧义消息），而不是 ToDictionary 抛 ArgumentException 崩溃。
    /// </summary>
    [Fact]
    public void Validate_ReportsAmbiguous_WhenTwoColumnDefinedNamesPointToSameCell()
    {
        using var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, SampleFillData(), MaterialsContract(), new SimpleExcelOptions { SheetName = "物料清单" });
        stream.Position = 0;

        using var conflicted = WithModifiedDefinedNames(stream, names =>
        {
            var code = names.Elements<DefinedName>().First(d => d.Name?.Value == "TF_Table_编码");
            var qty = names.Elements<DefinedName>().First(d => d.Name?.Value == "TF_Table_数量");
            qty.Text = code.Text; // "数量" 改指 "编码" 的同一单元格
        });

        var result = SimpleExcelContract.Validate(conflicted, MaterialsContract());

        Assert.False(result.IsValid);
        var ambiguous = result.Issues.Where(i => i.Code == TemplateValidationIssueCode.Ambiguous).ToList();
        Assert.Equal(2, ambiguous.Count);
        Assert.All(ambiguous, i => Assert.Equal("SimpleExcel.Contract.AmbiguousColumnPosition", i.MessageKey));
        Assert.Contains(ambiguous, i => i.Key == "编码");
        Assert.Contains(ambiguous, i => i.Key == "数量");

        // Read 同样不崩溃：歧义两列（编码/数量）都不参与定位、值补 null，其余列正常读
        conflicted.Position = 0;
        var loaded = SimpleExcelContract.Read(conflicted, MaterialsContract());
        var rows = loaded.Tables["Materials"];
        Assert.Equal(2, rows.Count);
        Assert.Null(rows[0]["编码"]);
        Assert.Equal("铝型材 6063-T5", rows[0]["名称"]);
        Assert.Null(rows[0]["数量"]);
    }

    [Fact]
    public void Validate_EnglishWrittenFile_ByDefinedNames_IsValid_WithoutLanguage()
    {
        using var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, SampleFillData(), MaterialsContract(), new SimpleExcelOptions { SheetName = "物料清单" }, En, EnLocalizer());
        stream.Position = 0;

        var result = SimpleExcelContract.Validate(stream, MaterialsContract());

        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(i => i.Message)));
    }

    [Fact]
    public void Service_FillEnglish_ThenParse_ReturnsStrongTypedData_WithoutLanguage()
    {
        var service = new EnMaterialsService();
        var data = new MaterialsData
        {
            Items =
            [
                new MaterialLine { Code = "AL-6063", Name = "铝型材 6063-T5", Qty = 120.5m, Date = new DateTime(2026, 8, 7), Enabled = true },
            ],
        };

        using var filled = service.Fill(data, new SimpleExcelOptions { SheetName = "物料清单" }, En, EnLocalizer());
        var parsed = service.Parse(filled, new SimpleExcelOptions { SheetName = "物料清单" });

        var item = Assert.Single(parsed.Items);
        Assert.Equal("AL-6063", item.Code);
        Assert.Equal("铝型材 6063-T5", item.Name);
        Assert.Equal(120.5m, item.Qty);
    }

    private static MemoryStream WithModifiedDefinedNames(Stream source, Action<DefinedNames> mutate)
    {
        source.Position = 0;
        using var buffer = new MemoryStream();
        source.CopyTo(buffer);
        buffer.Position = 0;

        using (var document = SpreadsheetDocument.Open(buffer, true))
        {
            if (document.WorkbookPart?.Workbook?.DefinedNames is { } names)
            {
                mutate(names);
            }

            document.Save();
        }

        buffer.Position = 0;
        var result = new MemoryStream();
        buffer.CopyTo(result);
        result.Position = 0;
        return result;
    }

    public sealed class EnMaterialsService : SimpleExcelTemplateService<MaterialsData>
    {
        protected override TemplateContract DefineContract() => MaterialsContract();
    }
}
