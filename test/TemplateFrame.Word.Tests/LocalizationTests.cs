using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Word.Tests;

/// <summary>国际化（迭代 12）：Word 插件校验消息按 CurrentUICulture 中英切换。</summary>
public sealed class LocalizationTests
{
    [Fact]
    public void Validate_ChineseCulture_UsesChineseMessages()
    {
        WithCulture("zh-CN", () =>
        {
            using var stream = TestDocuments.BuildDemoTemplate();
            var missing = ValidateMissingElement(stream);

            Assert.Equal("Word.Validation.MissingTextElement", missing.MessageKey);
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

            Assert.Equal("Word.Validation.MissingTextElement", missing.MessageKey);
            Assert.Contains("Missing the content control", missing.Message);
        });
    }

    private static TemplateValidationIssue ValidateMissingElement(System.IO.Stream stream)
    {
        var contract = TestDocuments.DemoContract() with
        {
            Elements = TestDocuments.DemoContract().Elements
                .Append(new TextElement { Key = "ExtraField", DisplayName = "新字段" })
                .ToArray(),
        };

        var result = new WordTemplateValidator().Validate(stream, contract);
        return Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing && i.Key == "ExtraField");
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