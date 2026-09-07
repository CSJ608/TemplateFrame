using System.Globalization;
using System.Threading;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.ConcurrencyTests;

// Linked into both format test projects; decorators retain real package creation, Save and Dispose.
public abstract class TemplateServiceConcurrencyTests
{
    protected abstract ITemplateEngine CreateEngine();
    protected abstract void Write(ITemplateBuilder builder, string value);
    protected abstract string Read(Stream stream);
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(15);

    [Theory]
    [InlineData("Create")]
    [InlineData("Layout")]
    [InlineData("Save")]
    [InlineData("Dispose")]
    public void SameInstance_SerializesWholeLifetime_AndIsolatesDocuments(string pauseAt)
    {
        using var paused = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var secondStarted = new ManualResetEventSlim();
        using var secondCreated = new ManualResetEventSlim();
        var engine = new TrackingEngine(CreateEngine());
        var service = new Service(engine, Write);
        engine.Creating = id => { if (id == 2) secondCreated.Set(); };
        engine.Stage = (builder, stage) =>
        {
            if (builder.Id == 1 && stage == pauseAt)
            {
                paused.Set();
                Assert.True(release.Wait(Timeout), "Test controller did not release the first call.");
            }
        };
        Stream? firstOutput = null, secondOutput = null;
        Exception? firstError = null, secondError = null;
        var first = new Thread(() =>
        {
            try { firstOutput = service.BuildInitialTemplateFile(CultureInfo.GetCultureInfo("zh-CN")); }
            catch (Exception e) { firstError = e; }
        })
        { IsBackground = true };
        var second = new Thread(() =>
        {
            secondStarted.Set();
            try { secondOutput = service.BuildInitialTemplateFile(CultureInfo.GetCultureInfo("en-US")); }
            catch (Exception e) { secondError = e; }
        })
        { IsBackground = true };
        first.Start();
        var secondWasStarted = false;
        try
        {
            Assert.True(paused.Wait(Timeout));
            second.Start();
            secondWasStarted = true;
            Assert.True(secondStarted.Wait(Timeout));
            // Dedicated thread has no other blocking operations before CreateBuilder.
            // Observe actual monitor contention, not an elapsed-time guess that it ran.
            Assert.True(SpinWait.SpinUntil(() => secondCreated.IsSet
                || (second.ThreadState & ThreadState.WaitSleepJoin) != 0
                || !second.IsAlive, Timeout));
            Assert.False(secondCreated.IsSet);
            Assert.True(second.IsAlive);
            Assert.Equal(ThreadState.WaitSleepJoin, second.ThreadState & ThreadState.WaitSleepJoin);
        }
        finally
        {
            // The controller releases first independently of second completing/entering the lock.
            release.Set();
            Assert.True(first.Join(Timeout));
            if (secondWasStarted) Assert.True(second.Join(Timeout));
        }
        using (firstOutput)
        using (secondOutput)
        {
            Assert.Null(firstError);
            Assert.Null(secondError);
            Assert.NotNull(firstOutput);
            Assert.NotNull(secondOutput);
            Assert.NotSame(firstOutput, secondOutput);
            Assert.Equal(0, firstOutput!.Position);
            Assert.Equal(0, secondOutput!.Position);
            Assert.Equal("1:zh-CN", Read(firstOutput));
            firstOutput.Dispose();
            Assert.True(secondOutput.CanRead);
            Assert.Equal("2:en-US", Read(secondOutput));
            Assert.False(service.HasBuilder);
            Assert.Equal(2, engine.Builders.Count);
            Assert.NotSame(engine.Builders[0].Inner, engine.Builders[1].Inner);
            Assert.All(engine.Builders, b =>
            {
                Assert.Equal(1, b.SaveCount);
                Assert.Equal(1, b.DisposeCount);
                Assert.True(b.Disposed);
            });
        }
    }

    [Theory]
    [InlineData("Create")]
    [InlineData("Layout")]
    [InlineData("Save")]
    [InlineData("Dispose")]
    public void Failure_ReleasesResources_AndNextGenerationRecovers(string failAt)
    {
        var engine = new TrackingEngine(CreateEngine());
        var service = new Service(engine, Write);
        var failure = new InvalidOperationException("injected " + failAt);
        engine.Stage = (builder, stage) =>
        {
            if (builder.Id == 1 && stage == failAt) throw failure;
        };
        Assert.Same(failure, Assert.Throws<InvalidOperationException>(() => service.BuildInitialTemplateFile()));
        Assert.False(service.HasBuilder);
        var failed = engine.Builders[0];
        Assert.Equal(1, failed.DisposeCount);
        Assert.True(failed.Disposed);
        if (failed.Target != null) Assert.False(failed.Target.CanWrite);
        // Recover on another thread so an accidentally retained reentrant lock cannot pass.
        Stream? recovered = null;
        Exception? recoveryError = null;
        var next = new Thread(() =>
        {
            try { recovered = service.BuildInitialTemplateFile(CultureInfo.GetCultureInfo("en-US")); }
            catch (Exception e) { recoveryError = e; }
        })
        { IsBackground = true };
        next.Start();
        Assert.True(next.Join(Timeout));
        using (recovered)
        {
            Assert.Null(recoveryError);
            Assert.NotNull(recovered);
            Assert.Equal("2:en-US", Read(recovered!));
        }
        Assert.Equal(1, engine.Builders[1].DisposeCount);
        Assert.False(service.HasBuilder);
    }

    [Fact]
    public void DifferentInstances_CanGenerateWhileAnotherInstanceIsPaused()
    {
        using var paused = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var engine = new TrackingEngine(CreateEngine());
        var service = new Service(engine, Write);
        engine.Stage = (_, stage) =>
        {
            if (stage != "Layout") return;
            paused.Set();
            Assert.True(release.Wait(Timeout));
        };
        Exception? error = null;
        var first = new Thread(() =>
        {
            try { using var output = service.BuildInitialTemplateFile(); }
            catch (Exception e) { error = e; }
        })
        { IsBackground = true };
        first.Start();
        try
        {
            Assert.True(paused.Wait(Timeout));
            var other = new Service(new TrackingEngine(CreateEngine()), Write);
            using var output = other.BuildInitialTemplateFile(CultureInfo.GetCultureInfo("en-US"));
            Assert.Equal("1:en-US", Read(output));
        }
        finally
        {
            release.Set();
            Assert.True(first.Join(Timeout));
        }
        Assert.Null(error);
    }

    [Fact]
    public void RecursiveLayout_DoesNotReplaceOuterBuilder()
    {
        var engine = new TrackingEngine(CreateEngine());
        var service = new Service(engine, Write);
        engine.Stage = (_, stage) =>
        {
            if (stage == "Layout")
                Assert.Throws<InvalidOperationException>(() => service.BuildInitialTemplateFile());
        };
        using var output = service.BuildInitialTemplateFile(CultureInfo.GetCultureInfo("en-US"));
        Assert.Equal("1:en-US", Read(output));
        Assert.Single(engine.Builders);
        Assert.False(service.HasBuilder);
    }

    private sealed class Service : TemplateService<object, TrackingBuilder>
    {
        private readonly Action<ITemplateBuilder, string> _write;
        public Service(TrackingEngine engine, Action<ITemplateBuilder, string> write) : base(engine) => _write = write;
        public bool HasBuilder => Builder != null;
        protected override TemplateContract DefineContract() => new();
        protected override void BuildInitialTemplate()
        {
            var initial = Builder;
            initial.Stage(initial, "Layout");
            Assert.Same(initial, Builder);
            Assert.False(initial.Disposed);
            _write(Builder.Inner, Builder.Id + ":" + Builder.Culture);
        }
    }

    private sealed class TrackingBuilder : ITemplateBuilder, IDisposable
    {
        public ITemplateBuilder Inner { get; }
        public int Id { get; }
        public string Culture { get; }
        public Action<TrackingBuilder, string> Stage { get; }
        public int SaveCount, DisposeCount;
        public bool Disposed;
        public Stream? Target;
        public TrackingBuilder(ITemplateBuilder inner, int id, string culture, Action<TrackingBuilder, string> stage)
        { Inner = inner; Id = id; Culture = culture; Stage = stage; }
        public void Save(Stream target)
        {
            Assert.False(Disposed);
            SaveCount++;
            Target = target;
            Stage(this, "Save");
            Inner.Save(target);
        }
        public void Dispose()
        {
            DisposeCount++;
            try { Stage(this, "Dispose"); }
            finally
            {
                ((IDisposable)Inner).Dispose();
                Disposed = true;
            }
        }
    }

    private sealed class TrackingEngine : ITemplateEngine
    {
        private readonly ITemplateEngine _inner;
        private int _nextId;
        public List<TrackingBuilder> Builders { get; } = new();
        public Action<TrackingBuilder, string> Stage = (_, _) => { };
        public Action<int> Creating = _ => { };
        public TrackingEngine(ITemplateEngine inner) => _inner = inner;
        public ITemplateBuilder CreateBuilder() => CreateBuilder(DefaultTemplateLocalizer.Instance, null);
        public ITemplateBuilder CreateBuilder(ITemplateLocalizer localizer, CultureInfo? culture)
        {
            var id = Interlocked.Increment(ref _nextId);
            Creating(id); // Signal entry before real format code can perform any blocking work.
            var builder = new TrackingBuilder(_inner.CreateBuilder(localizer, culture),
                id, culture?.Name ?? "zh-CN", (b, s) => Stage(b, s));
            lock (Builders) Builders.Add(builder);
            try { Stage(builder, "Create"); }
            catch { builder.Dispose(); throw; } // Factory owns resources until it returns.
            return builder;
        }
        public TemplateValidationResult Validate(Stream s, TemplateContract c) => _inner.Validate(s, c);
        public Stream Fill(Stream s, TemplateContract c, FillData d) => _inner.Fill(s, c, d);
        public TemplateFillResult FillDetailed(Stream s, TemplateContract c, FillData d) => _inner.FillDetailed(s, c, d);
        public FillData Parse(Stream s, TemplateContract c) => _inner.Parse(s, c);
        public TemplateParseResult ParseDetailed(Stream s, TemplateContract c) => _inner.ParseDetailed(s, c);
    }
}
