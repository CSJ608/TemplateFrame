using System.Globalization;
using TemplateFrame.Localization;
using Xunit;

namespace TemplateFrame.Tests;

/// <summary>
/// 迭代 13（文档内容 i18n）：<see cref="DefaultTemplateLocalizer"/>——
/// 占位符 / 页码默认文案按语言解析（zh 中性 + en 卫星）、查找顺序（业务注入 → 框架资源 → 键本身）、
/// 占位符一等语义（PlaceholderText / IsPlaceholderText，业务可覆盖 + 注册扩展）。
/// </summary>
public sealed class TemplateLocalizerTests
{
    private static readonly CultureInfo Zh = CultureInfo.GetCultureInfo("zh-CN");
    private static readonly CultureInfo En = CultureInfo.GetCultureInfo("en");

    [Fact]
    public void PlaceholderText_ChineseDefault_IsToBeFilledInChinese()
    {
        var localizer = DefaultTemplateLocalizer.Instance;

        Assert.Equal("待填充", localizer.PlaceholderText(Zh));
        Assert.Equal("待填充", localizer.PlaceholderText(CultureInfo.GetCultureInfo("zh")));
    }

    [Fact]
    public void PlaceholderText_English_IsToBeFilled()
    {
        var localizer = DefaultTemplateLocalizer.Instance;

        Assert.Equal("To be filled", localizer.PlaceholderText(En));
    }

    [Fact]
    public void PageNumberPattern_ChineseAndEnglish()
    {
        var localizer = DefaultTemplateLocalizer.Instance;

        Assert.Equal(
            "第{page}页，总{total}页",
            localizer.GetString(DefaultTemplateLocalizer.PageNumberPatternKey, Zh));
        Assert.Equal(
            "Page {page} of {total}",
            localizer.GetString(DefaultTemplateLocalizer.PageNumberPatternKey, En));
    }

    [Fact]
    public void GetString_MissingKey_ReturnsKeyItself()
    {
        var localizer = DefaultTemplateLocalizer.Instance;

        Assert.Equal("Doc.Missing", localizer.GetString("Doc.Missing", Zh));
    }

    [Fact]
    public void GetString_BusinessOverride_CultureSpecificThenNeutral()
    {
        var localizer = new DefaultTemplateLocalizer(new Dictionary<string, string>
        {
            ["en:Doc.Title"] = "Delivery Order",
            ["Doc.Title"] = "送货单",
        });

        Assert.Equal("送货单", localizer.GetString("Doc.Title", Zh));
        Assert.Equal("Delivery Order", localizer.GetString("Doc.Title", En));
    }

    [Fact]
    public void IsPlaceholderText_RecognizesBothLanguages_WithoutCultureDependency()
    {
        var localizer = DefaultTemplateLocalizer.Instance;

        Assert.True(localizer.IsPlaceholderText("待填充"));
        Assert.True(localizer.IsPlaceholderText("To be filled"));
        Assert.False(localizer.IsPlaceholderText(string.Empty));
        Assert.False(localizer.IsPlaceholderText("ABC"));
    }

    [Fact]
    public void PlaceholderText_BusinessOverride_AndExtraPlaceholders_AreRecognized()
    {
        var localizer = new DefaultTemplateLocalizer(
            new Dictionary<string, string> { [DefaultTemplateLocalizer.PlaceholderKey] = "请填写" },
            new[] { "待录入" });

        Assert.Equal("请填写", localizer.PlaceholderText(Zh));
        Assert.True(localizer.IsPlaceholderText("请填写"));   // 业务覆盖的占位符
        Assert.True(localizer.IsPlaceholderText("待录入"));   // 业务注册的扩展占位符
        Assert.True(localizer.IsPlaceholderText("待填充"));   // 框架默认仍识别（历史模板）
        Assert.True(localizer.IsPlaceholderText("To be filled"));
    }
}
