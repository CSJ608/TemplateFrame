using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;

// InProcess + 短迭代（2 预热 + 5 迭代）：文档生成类操作单次数十毫秒级，
// 牺牲一点统计强度换取整轮可跑完；正式发布对比时可去掉本配置用默认长跑。
var wordStudy = args.Contains("--word-study");
args = args.Where(arg => arg != "--word-study").ToArray();
var job = wordStudy
    ? Job.Default.WithLaunchCount(1).WithWarmupCount(4).WithIterationCount(8)
        .WithIterationTime(Perfolizer.Horology.TimeInterval.FromMilliseconds(200))
        .WithId("Word-Isolated")
    : Job.InProcess
        .WithWarmupCount(2)
        .WithIterationCount(5)
        .WithId("InProcess-Short");
var config = ManualConfig.Create(DefaultConfig.Instance).AddJob(job);

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);
