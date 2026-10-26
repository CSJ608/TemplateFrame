using System.Globalization;
using TemplateFrame.Localization;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class HeaderLocalizationOrderTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Headers_AreResolvedBeforePlaceholders_OnlyForKeyEntryPoint(bool localized)
    {
        var localizer = new RecordingLocalizer();
        using var builder = new ExcelTemplateBuilder(localizer, CultureInfo.GetCultureInfo("en"));
        if (localized)
            builder.AddTableKeys("Lines", ["A", "B"]);
        else
            builder.AddTable("Lines", ["A", "B"]);

        Assert.Equal(localized
            ? new[] { "header:A:en", "header:B:en", "placeholder:en", "placeholder:en" }
            : new[] { "placeholder:en", "placeholder:en" }, localizer.Calls);
    }

    [Fact]
    public void InvalidColumns_AreRejectedBeforeLocalization()
    {
        var localizer = new RecordingLocalizer();
        using var builder = new ExcelTemplateBuilder(localizer, CultureInfo.GetCultureInfo("en"));

        Assert.Throws<ArgumentException>(() => builder.AddTableKeys("Lines", []));
        Assert.Throws<ArgumentNullException>(() => builder.AddTableKeys("Lines", null!));
        Assert.Empty(localizer.Calls);
    }

    [Fact]
    public void HeaderFailure_PropagatesSameException_BeforeAnyPlaceholder()
    {
        var failure = new InvalidOperationException("localizer failure");
        var localizer = new RecordingLocalizer { Failure = failure };
        using var builder = new ExcelTemplateBuilder(localizer, CultureInfo.GetCultureInfo("en"));

        var actual = Assert.Throws<InvalidOperationException>(() => builder.AddTableKeys("Lines", ["A", "B", "C"]));

        Assert.Same(failure, actual);
        Assert.Equal(new[] { "header:A:en", "header:B:en" }, localizer.Calls);
    }

    private sealed class RecordingLocalizer : ITemplateLocalizer
    {
        public List<string> Calls { get; } = [];
        public Exception? Failure { get; init; }

        public string GetString(string key, CultureInfo? culture = null)
        {
            Calls.Add("header:" + key + ":" + culture?.Name);
            if (key == "B" && Failure is { } failure)
                throw failure;
            return "translated-" + key;
        }

        public string PlaceholderText(CultureInfo? culture = null)
        {
            Calls.Add("placeholder:" + culture?.Name);
            return "pending";
        }

        public bool IsPlaceholderText(string text) => text == "pending";
    }
}
