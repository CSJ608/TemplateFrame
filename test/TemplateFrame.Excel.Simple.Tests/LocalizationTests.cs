using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

/// <summary>国际化（迭代 12）：SimpleExcel 插件校验消息按 CurrentUICulture 中英切换。</summary>
public sealed class LocalizationTests
{
    [Fact]
    public void Validate_ChineseCulture_UsesChineseMessages()
    {
        WithCulture("zh-CN", () =>
        {
            using var stream = BuildStreamMissingColumn();
            var missing = ValidateMissingColumn(stream);

            Assert.Equal("SimpleExcel.Contract.MissingColumn", missing.MessageKey);
            Assert.Contains("缺少列", missing.Message);
        });
    }

    [Fact]
    public void Validate_EnglishCulture_UsesEnglishMessages()
    {
        WithCulture("en", () =>
        {
            using var stream = BuildStreamMissingColumn();
            var missing = ValidateMissingColumn(stream);

            Assert.Equal("SimpleExcel.Contract.MissingColumn", missing.MessageKey);
            Assert.Contains("missing column", missing.Message);
        });
    }

    private static MemoryStream BuildStreamMissingColumn()
    {
        var table = new SimpleExcelTable { Headers = ["编码", "数量"], Rows = [["AL-6063", 120.5]] };
        var stream = new MemoryStream();
        SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "物料清单" });
        stream.Position = 0;
        return stream;
    }

    private static TemplateValidationIssue ValidateMissingColumn(System.IO.Stream stream)
    {
        var contract = new TemplateContract
        {
            Name = "Materials",
            Elements =
            [
                new TableElement
                {
                    Key = "Materials",
                    DisplayName = "物料清单",
                    Columns =
                    [
                        new TextElement { Key = "编码", DisplayName = "编码", Required = true },
                        new TextElement { Key = "名称", DisplayName = "名称", Required = true },
                    ],
                },
            ],
        };

        var result = SimpleExcelContract.Validate(stream, contract);
        return Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing && i.Key == "名称");
    }

    private static void WithCulture(string name, Action action)
    {
        var original = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentUICulture = new CultureInfo(name);
            action();
        }
        finally
        {
            CultureInfo.CurrentUICulture = original;
        }
    }
}
