using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using TemplateFrame.CapacityProbe;

var settings = ProbeSettings.Parse(args);
Directory.CreateDirectory(settings.Directory);
Emit(new
{
    Phase = "prepare",
    settings,
    Runtime = RuntimeInformation.FrameworkDescription,
    ProcessId = Environment.ProcessId,
    HeapBudgetBytes = GC.GetGCMemoryInfo().TotalAvailableMemoryBytes,
    ServerGc = System.Runtime.GCSettings.IsServerGC,
});
try
{
    if (settings.Mode == "shared-profile-read")
    {
        SharedStringProfile.Run(settings);
        return;
    }
    if (settings.Mode is "generate" or "generate-shared")
    {
        ExcelFixture.Write(settings);
        Emit(new { Phase = "complete", Status = "generated" });
        return;
    }

    var work = Enumerable.Range(0, settings.Concurrency).Select(i => new Workload(settings, i)).ToArray();
    using var ready = new CountdownEvent(work.Length);
    using var start = new ManualResetEventSlim();
    var tasks = work.Select(item => Task.Factory.StartNew(() =>
    {
        ready.Signal();
        start.Wait();
        return item.Run();
    }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default)).ToArray();
    ready.Wait();
    Emit(new { Phase = "operation", ManagedBytes = GC.GetTotalMemory(false) });
    var timer = Stopwatch.StartNew();
    start.Set();
    var durations = await Task.WhenAll(tasks);
    timer.Stop();
    Array.Sort(durations);
    Emit(new
    {
        Phase = "validation",
        BatchMilliseconds = timer.Elapsed.TotalMilliseconds,
        MinimumMilliseconds = durations[0],
        MedianMilliseconds = durations[durations.Length / 2],
        MaximumMilliseconds = durations[^1],
        ManagedBytes = GC.GetTotalMemory(false),
    });
    foreach (var item in work) item.Validate();
    Emit(new
    {
        Phase = "complete",
        Status = "passed",
        OutputBytes = work.Sum(item => item.OutputBytes),
        PeakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64,
        ScratchPeakBytes = work.Sum(item => item.ScratchPeakBytes),
        ScratchRemainingFiles = work.Sum(item => item.ScratchRemainingFiles)
    });
}
catch (Exception exception)
{
    Emit(new { Phase = "complete", Status = "failed", Error = exception.ToString() });
    Environment.ExitCode = 1;
}

static void Emit(object value)
{
    Console.WriteLine(JsonSerializer.Serialize(value));
    Console.Out.Flush();
}
