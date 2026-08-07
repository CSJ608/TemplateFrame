using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Mapping;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Tests;

/// <summary>
/// 国际化（迭代 12）：运行时消息按 <see cref="CultureInfo.CurrentUICulture"/> 中英切换——中文为中性文化默认，
/// en 卫星资源命中时返回英文。校验消息还暴露 MessageKey + MessageArgs 供调用方自行本地化。
/// </summary>
public sealed class LocalizationTests
{
    public sealed class SampleData
    {
        public string No { get; set; } = "DO001";

        public DateTime OrderDate { get; set; }
    }

    [Fact]
    public void ValidationMessage_ChineseCulture_IsChinese()
    {
        WithCulture("zh-CN", () =>
        {
            var result = ValidateMissingField();

            var issue = Assert.Single(result.Issues);
            Assert.Equal(TemplateValidationIssueCode.Missing, issue.Code);
            Assert.Equal("Validation.DataMissingRequiredField", issue.MessageKey);
            Assert.Contains("数据缺少必填字段", issue.Message);
        });
    }

    [Fact]
    public void ValidationMessage_EnglishCulture_IsEnglish()
    {
        WithCulture("en", () =>
        {
            var result = ValidateMissingField();

            var issue = Assert.Single(result.Issues);
            Assert.Equal("Validation.DataMissingRequiredField", issue.MessageKey);
            Assert.Contains("missing required field", issue.Message);
        });
    }

    [Fact]
    public void ValidationIssue_ExposesMessageKeyAndArgs()
    {
        var result = ValidateMissingField();

        var issue = Assert.Single(result.Issues);
        Assert.Equal("Validation.DataMissingRequiredField", issue.MessageKey);
        Assert.Equal(new object?[] { "No", "单号" }, issue.MessageArgs);
    }

    [Fact]
    public void MappingException_ChineseCulture_IsChinese()
    {
        WithCulture("zh-CN", () =>
        {
            var ex = Assert.Throws<InvalidOperationException>(
                () => DataPathMapper.ToFillData(new SampleData(), DuplicatePathContract()));

            Assert.Contains("DataPath 重复", ex.Message);
        });
    }

    [Fact]
    public void MappingException_EnglishCulture_IsEnglish()
    {
        WithCulture("en", () =>
        {
            var ex = Assert.Throws<InvalidOperationException>(
                () => DataPathMapper.ToFillData(new SampleData(), DuplicatePathContract()));

            Assert.Contains("duplicate DataPath", ex.Message);
        });
    }

    private static TemplateValidationResult ValidateMissingField()
    {
        var contract = new TemplateContract
        {
            Name = "Localization",
            Elements =
            [
                new TextElement { Key = "No", DisplayName = "单号", DataPath = "No" },
            ],
        };

        return new TemplateDataValidator().Validate(new FillData(), contract);
    }

    private static TemplateContract DuplicatePathContract()
        => new()
        {
            Name = "Dup",
            Elements =
            [
                new TextElement { Key = "X", DisplayName = "X", DataPath = "No" },
                new TextElement { Key = "Y", DisplayName = "Y", DataPath = "No" },
            ],
        };

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