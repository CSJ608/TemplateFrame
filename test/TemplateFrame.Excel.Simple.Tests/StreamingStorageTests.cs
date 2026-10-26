using TemplateFrame.Engine;
using TemplateFrame.Internal.Excel;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class StreamingStorageTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(2)]
    public void DiskStrings_AppendingAfterRandomReadsPreservesOffsetsAndNulls(int cacheEntries)
    {
        using var fixture = new ScratchDirectory();
        var options = new SimpleExcelStreamingReadOptions { TemporaryDirectory = fixture.Path, SharedStringCacheEntries = cacheEntries };
        using (var storage = new ExcelTemporaryStorage(options))
        {
            var strings = new ExcelDiskStrings(storage, options);
            var expected = new List<string?>();
            long bytes = 0;
            for (var index = 0; index < 5000; index++)
            {
                string? value = index % 3 == 0 ? null : index % 3 == 1 ? "" : "中文😀" + index;
                strings.AddNullable(value);
                expected.Add(value);
                bytes += 12 + (value is null ? 0 : System.Text.Encoding.UTF8.GetByteCount(value));
                Assert.Equal(expected[index / 2], strings.ReadNullable(index / 2));
                Assert.Equal(expected[0], strings.ReadNullable(0));
            }
            for (var index = expected.Count - 1; index >= 0; index--)
                Assert.Equal(expected[index], strings.ReadNullable(index));
            Assert.Equal(bytes, storage.ReservedBytes);
        }
        Assert.Empty(Directory.GetFiles(fixture.Path));
    }

    [Fact]
    public void DiskStrings_PreservesNullSeparatelyFromEmptyText()
    {
        using var fixture = new ScratchDirectory();
        var options = new SimpleExcelStreamingReadOptions { TemporaryDirectory = fixture.Path };
        using (var storage = new ExcelTemporaryStorage(options))
        {
            var strings = new ExcelDiskStrings(storage, options);
            strings.AddNullable(null);
            strings.Add("");
            strings.Add("中");
            for (var repeat = 0; repeat < 3; repeat++)
            {
                Assert.Equal("中", strings.ReadNullable(2));
                Assert.Equal("", strings.ReadNullable(1));
                Assert.Null(strings.ReadNullable(0));
            }
            Assert.Equal(39, storage.ReservedBytes);
        }
        Assert.Empty(Directory.GetFiles(fixture.Path));
    }

    [Fact]
    public void DiskStrings_RandomReadsUseBoundedCacheAndCleanupFiles()
    {
        using var fixture = new ScratchDirectory();
        var options = new SimpleExcelStreamingReadOptions
        {
            TemporaryDirectory = fixture.Path,
            SharedStringCacheCharacters = 12,
            SharedStringCacheEntries = 2,
        };
        using (var storage = new ExcelTemporaryStorage(options))
        {
            var strings = new ExcelDiskStrings(storage, options);
            for (var index = 0; index < 1000; index++) strings.Add("中文" + index);
            for (var index = 999; index >= 0; index--)
            {
                Assert.Equal("中文" + index, strings[index]);
                Assert.InRange(strings.CachedCharacters, 0, 12);
                Assert.InRange(strings.CachedEntries, 0, 2);
            }
            Assert.Equal(1000, strings.Count);
            Assert.Equal(2, Directory.GetFiles(fixture.Path).Length);
            Assert.Throws<ArgumentOutOfRangeException>(() => strings[-1]);
            Assert.Throws<ArgumentOutOfRangeException>(() => strings[1000]);
        }
        Assert.Empty(Directory.GetFiles(fixture.Path));
    }

    [Fact]
    public void DiskStrings_ChargesIndexAndUtf8BytesBeforeWriting()
    {
        using var fixture = new ScratchDirectory();
        var options = new SimpleExcelStreamingReadOptions { TemporaryDirectory = fixture.Path, MaxTemporaryBytes = 15 };
        using (var storage = new ExcelTemporaryStorage(options))
        {
            var strings = new ExcelDiskStrings(storage, options);
            strings.Add("中");
            Assert.Equal(15, storage.ReservedBytes);
            Assert.Throws<InvalidOperationException>(() => strings.Add(""));
            Assert.Single(strings);
            Assert.Equal("中", strings[0]);
        }
        Assert.Empty(Directory.GetFiles(fixture.Path));
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 1)]
    public void DiskStrings_LargeValuesBypassCacheAndEmptyValuesStayBounded(int characters, int entries)
    {
        using var fixture = new ScratchDirectory();
        var options = new SimpleExcelStreamingReadOptions
        {
            TemporaryDirectory = fixture.Path,
            SharedStringCacheCharacters = characters,
            SharedStringCacheEntries = entries,
        };
        using var storage = new ExcelTemporaryStorage(options);
        var strings = new ExcelDiskStrings(storage, options);
        strings.Add(new string('x', 32767));
        for (var index = 0; index < 100; index++) strings.Add("");
        Assert.Equal(32767, strings[0].Length);
        for (var index = 1; index <= 100; index++) Assert.Empty(strings[index]);
        Assert.InRange(strings.CachedEntries, 0, entries);
        Assert.Throws<InvalidOperationException>(() => strings.Add(new string('x', 32768)));
    }

    [Theory]
    [InlineData("input")]
    [InlineData("disk")]
    [InlineData("cancel")]
    public void BufferInput_LimitsAndCancellationLeaveSourceOpenAndRemoveScratch(string failure)
    {
        using var fixture = new ScratchDirectory();
        using var source = new MemoryStream(new byte[100]);
        var options = new SimpleExcelStreamingReadOptions
        {
            TemporaryDirectory = fixture.Path,
            MaxTemporaryBytes = failure == "disk" ? 50 : 1000,
            ResourceLimits = new TemplateResourceLimits { MaxInputBytes = failure == "input" ? 50 : 1000 },
        };
        using (var storage = new ExcelTemporaryStorage(options))
        {
            if (failure == "cancel") Assert.Throws<OperationCanceledException>(() => storage.BufferInput(source, new CancellationToken(true)));
            else Assert.Throws<InvalidOperationException>(() => storage.BufferInput(source, default));
        }
        Assert.True(source.CanRead);
        Assert.Empty(Directory.GetFiles(fixture.Path));
    }

    [Fact]
    public void BufferInput_UsesSameBudgetAsStrings()
    {
        using var fixture = new ScratchDirectory();
        var options = new SimpleExcelStreamingReadOptions { TemporaryDirectory = fixture.Path, MaxTemporaryBytes = 20 };
        using var storage = new ExcelTemporaryStorage(options);
        using var source = new MemoryStream(new byte[10]);
        var buffered = storage.BufferInput(source, default);
        Assert.Equal(0, buffered.Position);
        Assert.Equal(10, buffered.Length);
        Assert.Equal(10, storage.ReservedBytes);
        var strings = new ExcelDiskStrings(storage, options);
        Assert.Throws<InvalidOperationException>(() => strings.Add(""));
        Assert.True(source.CanRead);
    }

    private sealed class ScratchDirectory : IDisposable
    {
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "TemplateFrame-tests-" + Guid.NewGuid().ToString("N"));
        internal ScratchDirectory() => Directory.CreateDirectory(Path);
        public void Dispose() => Directory.Delete(Path, recursive: false);
    }
}
