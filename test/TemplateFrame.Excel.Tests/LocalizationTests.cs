using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Tests;

/// <summary>国际化（迭代 12）：Excel 插件校验消息按 CurrentUICulture 中英切换。</summary>
public sealed class LocalizationTests
{
    [Fact]
    public void Validate_ChineseCulture_UsesChineseMessages()
    {
        WithCulture("zh-CN", () =>
        {
            using var stream = TestDocuments.BuildDemoTemplate();
            var missing = ValidateMissingElement(stream);

            Assert.Equal("Excel.Validation.MissingTextElement", missing.MessageKey);
            Assert.Contains("缺少文本元素", missing.Message);
        });
    }

    [Fact]
    public void Validate_EnglishCulture_UsesEnglishMessages()
    {
        WithCulture("en", () =>
        {
            using var stream = TestDocuments.BuildDemoTemplate();
            var missing = ValidateMissingElement(stream);

            Assert.Equal("Excel.Validation.MissingTextElement", missing.MessageKey);
            Assert.Contains("Missing the defined name", missing.Message);
        });
    }

    private static TemplateValidationIssue ValidateMissingElement(System.IO.Stream stream)
    {
        var contract = TestDocuments.DemoContract() with
        {
            Elements = TestDocuments.DemoContract().Elements
                .Append(new TextElement { Key = "MissingField", DisplayName = "缺失字段" })
                .ToList(),
        };

        var result = new ExcelTemplateValidator().Validate(stream, contract);
        return Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing && i.Key == "MissingField");
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