using System.Reflection;
using System.Runtime.CompilerServices;
using TemplateFrame.Contract;
using TemplateFrame.Mapping;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class DataPathMapperCacheTests
{
    // Inspect identity without adding a public cache API or depending on cache container internals.
    private static object Mapping(TemplateContract contract, Type type)
        => typeof(DataPathMapper).GetMethod("GetMapping", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, new object[] { contract, type })!;

    private static TemplateContract Contract() => new()
    {
        Name = "CacheLifetime",
        Elements = [new TextElement { Key = "value", DataPath = "Value" }],
    };

    public sealed class NumberDto
    {
        public int Value { get; set; }
    }

    public sealed class TextDto
    {
        public string Value { get; set; } = "";
    }

    [Fact]
    public void ReleasedContracts_ElementsAndMappingsAreCollectible()
    {
        var references = CreateReleasedMappings();
        Collect();
        Assert.All(references, reference => Assert.False(reference.IsAlive));
    }

    // GC runs only after this frame returns, in both Debug and optimized builds.
    // Return no strong references, including to the last loop iteration or mapping.
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static WeakReference[] CreateReleasedMappings()
    {
        var references = new List<WeakReference>();
        for (var i = 0; i < 100; i++)
        {
            var contract = Contract();
            var fill = DataPathMapper.ToFillData(new NumberDto { Value = i }, contract);
            Assert.Equal(i, DataPathMapper.FromFillData<NumberDto>(fill, contract).Value);
            references.Add(new WeakReference(contract));
            references.Add(new WeakReference(contract.Elements));
            references.Add(new WeakReference(contract.Elements[0]));
            references.Add(new WeakReference(Mapping(contract, typeof(NumberDto))));
        }
        return references.ToArray();
    }

    private static void Collect()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    [Fact]
    public void ActiveContract_ReusesMappingAcrossCollectionsAndBothDirections()
    {
        var contract = Contract();
        var reference = CreateWeakMapping(contract);
        Collect();
        var original = reference.Target;
        Assert.NotNull(original);
        var fill = DataPathMapper.ToFillData(new NumberDto { Value = 42 }, contract);
        Assert.Equal(42, DataPathMapper.FromFillData<NumberDto>(fill, contract).Value);
        Assert.Same(original, Mapping(contract, typeof(NumberDto)));
        GC.KeepAlive(contract);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static WeakReference CreateWeakMapping(TemplateContract contract)
        => new(Mapping(contract, typeof(NumberDto)));

    [Fact]
    public void EqualButDistinctContracts_HaveIndependentMappings()
    {
        var contract = Contract();
        var equalContract = contract with { };
        Assert.Equal(contract, equalContract);
        Assert.NotSame(contract, equalContract);
        Assert.NotSame(Mapping(contract, typeof(NumberDto)), Mapping(equalContract, typeof(NumberDto)));
    }

    [Fact]
    public void SameContract_IsolatesDtoTypes()
    {
        var contract = Contract();
        var fill = DataPathMapper.ToFillData(new TextDto { Value = "123" }, contract);
        Assert.Equal(123, DataPathMapper.FromFillData<NumberDto>(fill, contract).Value);
        Assert.Equal("123", DataPathMapper.FromFillData<TextDto>(fill, contract).Value);
        var numeric = DataPathMapper.ToFillData(new NumberDto { Value = 456 }, contract);
        Assert.Equal("456", DataPathMapper.FromFillData<TextDto>(numeric, contract).Value);
        Assert.NotSame(Mapping(contract, typeof(NumberDto)), Mapping(contract, typeof(TextDto)));
    }

    [Fact]
    public async Task ConcurrentFirstAccess_PublishesOneMappingPerTypeAndIsolatesResults()
    {
        var contract = Contract();
        using var start = new ManualResetEventSlim();
        using var ready = new CountdownEvent(8);
        var workers = Enumerable.Range(0, 8).Select(worker => Task.Factory.StartNew(() =>
        {
            ready.Signal();
            Assert.True(start.Wait(TimeSpan.FromSeconds(15)));
            var numberMapping = Mapping(contract, typeof(NumberDto));
            var textMapping = Mapping(contract, typeof(TextDto));
            for (var i = 0; i < 50; i++)
            {
                var value = worker * 50 + i;
                var fill = DataPathMapper.ToFillData(new NumberDto { Value = value }, contract);
                Assert.Equal(value, DataPathMapper.FromFillData<NumberDto>(fill, contract).Value);
                Assert.Equal(value.ToString(), DataPathMapper.FromFillData<TextDto>(fill, contract).Value);
                Assert.Same(numberMapping, Mapping(contract, typeof(NumberDto)));
                Assert.Same(textMapping, Mapping(contract, typeof(TextDto)));
            }
            return (numberMapping, textMapping);
        }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default)).ToArray();
        bool allReady;
        try
        {
            allReady = ready.Wait(TimeSpan.FromSeconds(15));
        }
        finally
        {
            start.Set();
        }
        var results = await Task.WhenAll(workers);
        Assert.True(allReady);
        Assert.All(results, result =>
        {
            Assert.Same(results[0].numberMapping, result.numberMapping);
            Assert.Same(results[0].textMapping, result.textMapping);
            Assert.NotSame(result.numberMapping, result.textMapping);
        });
    }
}
