using System.IO.Compression;
using TemplateFrame.Engine;
using TemplateFrame.Internal;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class ResourceLimitTests
{
    [Fact]
    public void PackageLimits_RejectExpandedDataAndEntryCounts_AndRestorePosition()
    {
        using var package = CreatePackage(2, 1024);
        package.Position = 7;
        var limits = TemplateResourceLimits.Default with { MaxExpandedBytes = 1500 };
        Assert.Throws<InvalidOperationException>(() => StreamUtil.ValidatePackage(package, limits));
        Assert.Equal(7, package.Position);
        Assert.True(package.CanRead);

        Assert.Throws<InvalidOperationException>(() => StreamUtil.ValidatePackage(package,
            TemplateResourceLimits.Default with { MaxPackageEntries = 1 }));
        Assert.Throws<InvalidOperationException>(() => StreamUtil.ValidatePackage(package,
            TemplateResourceLimits.Default with { MaxPackageEntryBytes = 1000 }));
    }

    [Fact]
    public void PackageLimits_AcceptExactBoundaryWithoutTakingOwnership()
    {
        using var package = CreatePackage(2, 1024);
        var limits = TemplateResourceLimits.Default with
        {
            MaxInputBytes = package.Length,
            MaxPackageEntries = 2,
            MaxPackageEntryBytes = 1024,
            MaxExpandedBytes = 2048,
        };
        package.Position = 3;
        StreamUtil.ValidatePackage(package, limits);
        Assert.Equal(3, package.Position);
        var bytes = StreamUtil.ReadAllBytes(package, limits);
        Assert.Equal(package.Length, bytes.LongLength);
        Assert.True(package.CanRead);
    }

    [Fact]
    public void InputLimit_AppliesToNonSeekableStreams()
    {
        using var package = CreatePackage(1, 128);
        using var source = new NonSeekableStream(package);
        var limits = TemplateResourceLimits.Default with { MaxInputBytes = package.Length - 1 };
        Assert.Throws<InvalidOperationException>(() => StreamUtil.ReadAllBytes(source, limits));
        Assert.True(source.CanRead);
    }

    [Fact]
    public void Images_AreBoundedForBothArraysAndStreams()
    {
        var limits = TemplateResourceLimits.Default with { MaxImageBytes = 4 };
        var bytes = new byte[4];
        Assert.Same(bytes, StreamUtil.ToBytes(bytes, limits));
        Assert.Throws<InvalidOperationException>(() => StreamUtil.ToBytes(new byte[5], limits));
        using var stream = new MemoryStream(new byte[5]);
        Assert.Throws<InvalidOperationException>(() => StreamUtil.ToBytes(stream, limits));
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void TableBudget_IsCumulativeAndCannotOverflow()
    {
        var limits = TemplateResourceLimits.Default with { MaxTableRows = 3, MaxTableCells = 7 };
        var budget = new TemplateResourceBudget(limits);
        budget.AddCells(1);
        budget.AddRows(2, 2);
        Assert.Throws<InvalidOperationException>(() => budget.AddRows(1, 3));
        budget.AddRows(1, 2);
        Assert.Throws<InvalidOperationException>(() => budget.AddRows(1, 0));
        var huge = new TemplateResourceBudget(TemplateResourceLimits.Default with { MaxTableCells = long.MaxValue });
        Assert.Throws<InvalidOperationException>(() => huge.AddRows(2, long.MaxValue));
    }

    [Fact]
    public void CloneBudget_IsCumulativeAndCannotOverflow()
    {
        var budget = new TemplateResourceBudget(TemplateResourceLimits.Default with { MaxClonedXmlCharacters = 12 });
        budget.AddClonedXml(3, 2);
        budget.AddClonedXml(2, 3);
        Assert.Throws<InvalidOperationException>(() => budget.AddClonedXml(1, 1));
        var huge = new TemplateResourceBudget(TemplateResourceLimits.Default with { MaxClonedXmlCharacters = long.MaxValue });
        Assert.Throws<InvalidOperationException>(() => huge.AddClonedXml(long.MaxValue, 2));
    }

    [Fact]
    public void InvalidConfiguration_IsRejectedBeforeReading()
    {
        using var stream = new MemoryStream();
        Assert.Throws<ArgumentOutOfRangeException>(() => StreamUtil.ReadAllBytes(stream,
            TemplateResourceLimits.Default with { MaxXmlCharactersPerPart = 0 }));
    }

    [Fact]
    public void PackageLimits_EntryLimitPrecedesTotalLimit_AndRestoresPosition()
    {
        using var package = CreatePackage(1, 1024);
        package.Position = 5;

        var error = Assert.Throws<InvalidOperationException>(() => StreamUtil.ValidatePackage(package,
            TemplateResourceLimits.Default with { MaxPackageEntryBytes = 100, MaxExpandedBytes = 50 }));

        Assert.Contains(nameof(TemplateResourceLimits.MaxPackageEntryBytes), error.Message);
        Assert.Equal(5, package.Position);
        Assert.True(package.CanRead);
    }

    [Fact]
    public void InvalidPackage_WrapsInvalidData_AndRestoresPosition()
    {
        using var package = new MemoryStream(new byte[32]);
        package.Position = 7;

        var error = Assert.Throws<InvalidOperationException>(() => StreamUtil.ValidatePackage(package, TemplateResourceLimits.Default));

        Assert.IsType<InvalidDataException>(error.InnerException);
        Assert.Equal(7, package.Position);
        Assert.True(package.CanRead);
    }

    private static MemoryStream CreatePackage(int entries, int expandedBytes)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            for (var i = 0; i < entries; i++)
            {
                using var output = archive.CreateEntry($"part{i}.bin").Open();
                var bytes = new byte[expandedBytes];
                output.Write(bytes, 0, bytes.Length);
            }
        }
        stream.Position = 0;
        return stream;
    }

    private sealed class NonSeekableStream(Stream inner) : Stream
    {
        public override bool CanRead => inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
