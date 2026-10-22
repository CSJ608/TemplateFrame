# 从 2.x 迁移到 3.0

3.0 默认启用有限资源预算。合法的大文档或表格在 2.x 中可能成功处理，升级后可能因超限而失败；这一默认行为变化是本次提升主版本的原因。例如，从 A1 写入 100,001 行单列表格仍符合 Excel 的行列范围，但会超过默认的 100,000 数据行预算。

原有公开入口签名保留，四个库仍提供 `netstandard2.0 / net462 / net8.0` 资产。普通文档可以继续使用原调用方式；大文档业务应先确认容量，再显式传入 `TemplateResourceLimits`，并让同一业务的读取、校验与填充采用一致配置。建议将使用到的 TemplateFrame 包一并升级到 3.0.0。

## 选择业务预算

完整默认值与作用范围见 [安全说明的默认资源限额](SECURITY.md#默认资源限额)。配置只覆盖业务需要调整的属性，其余保留默认值：

- `MaxInputBytes`、`MaxPackageEntryBytes`、`MaxExpandedBytes`、`MaxImageBytes` 以字节计；1 MiB = 1,048,576 字节。
- `MaxXmlCharactersPerPart` 和 `MaxClonedXmlCharacters` 以 XML 字符数计，分别限制读取的单个 XML part、一次填充中新增行副本的累计 XML。模板行中的静态内容也会消耗克隆预算。
- `MaxPackageEntries` 计包内条目数；`MaxTableRows` 和 `MaxTableCells` 按单次操作累计，多个表格共用该次预算。Simple 的单元格预算也计入表头；读取时访问的全空物理数据行仍可能消耗预算。

所有值必须大于零，`0` 不表示无限制，无效配置会抛 `ArgumentOutOfRangeException`。超限的 Read / Parse / Fill / Write 抛 `InvalidOperationException`，Validate 将输入资源问题报告为 `Invalid`；不会通过截断数据返回伪成功。提高行数预算不会同时提高 XML、单元格或包大小预算，也不能放宽 Excel 的合法坐标范围。

每次操作独立计数，重用引擎不会把上次消耗带入下次。预算不等于峰值内存或超时保证；按真实版式、输入规模和业务并发验证配置，不要统一设置为类型最大值。

## Word 与 Excel

通过 `TemplateFillOptions.ResourceLimits` 配置引擎，限额会同时用于 Validate、Fill 与 Parse。下面的数值是业务配置示例，并非所有项目都应采用的推荐容量。

```csharp
using TemplateFrame.Engine;
using TemplateFrame.Word;
using TemplateFrame.Excel;

var limits = TemplateResourceLimits.Default with
{
    MaxTableRows = 20_000,
    MaxTableCells = 200_000,
    MaxClonedXmlCharacters = 32L * 1024 * 1024,
};

var options = new TemplateFillOptions { ResourceLimits = limits };
var wordEngine = new WordTemplateEngine(options);
var excelEngine = new ExcelTemplateEngine(options);

// 独立 Parser 的新增构造器需要同时传入 limits 和 localizer。
var wordParser = new WordTemplateParser(limits, localizer: null);
var excelParser = new ExcelTemplateParser(limits, localizer: null);
var wordValidator = new WordTemplateValidator(limits);
var excelValidator = new ExcelTemplateValidator(limits);
```

使用自定义 `ITemplateLocalizer` 时继续传入原实例。Validate 检查通过不保证随后提供的业务数据一定在填充预算内；尤其 Word 的结构校验不会预先检查尚未传入的填充行数。

## Excel.Simple：100,001 行迁移示例

以下示例显式允许最多 120,000 个数据行、150,000 个处理单元格，其余预算保持默认。默认写入调用会拒绝这张 100,001 行的合法单列表格；配置后的写入和直接读取使用相同限额。

```csharp
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TemplateFrame.Engine;
using TemplateFrame.Excel.Simple;

var table = new SimpleExcelTable
{
    Headers = new[] { "Value" },
    Rows = Enumerable.Range(1, 100_001)
        .Select(i => (IReadOnlyList<object?>)new object?[] { (long)i })
        .ToArray(),
};
var limits = TemplateResourceLimits.Default with
{
    MaxTableRows = 120_000,
    MaxTableCells = 150_000,
};
var options = new SimpleExcelOptions { ResourceLimits = limits };

using var output = new MemoryStream();
// SimpleExcel.Write(output, table); // 3.0 默认抛 InvalidOperationException。
SimpleExcel.Write(output, table, options);
output.Position = 0;
var restored = SimpleExcel.Read(output, tableName: null, resourceLimits: limits);
```

直接 `SimpleExcel.Read` 使用三参数重载；契约入口与强类型服务则接收 `SimpleExcelOptions`。下表中的流、契约、数据与服务是业务已有对象，`options` 为上述配置：

| 入口 | 配置方式 |
|---|---|
| 契约写入 | `SimpleExcelContract.Write(target, fillData, contract, options)` |
| 契约读取 | `SimpleExcelContract.Read(source, contract, options)` |
| 契约校验 | `SimpleExcelContract.Validate(source, contract, options)` |
| `SimpleExcelTemplateService<TData>` 模板生成 | `service.BuildTemplate(options)` |
| 强类型服务填充 | `service.Fill(data, options)` |
| 强类型服务读取 | `service.Parse(source, options)` |
| 强类型服务校验 | `service.Validate(source, options)` |

服务的限额通过每次调用的 options 传递。连续操作同一可定位流时按接口约定归零；调用方负责释放自己提供的流，以及 Fill / BuildTemplate 返回的流。

## 其他升级注意事项

`TemplateFillOptions` 与 `SimpleExcelOptions` 是 record，新增的 `ResourceLimits` 属性会参与相等比较，并可能改变默认序列化结果的字段集合。依赖配置比较、快照或固定 JSON 结构的业务应同步检查。

资源预算不进行宏、OLE 或外部关系净化，Fill 仍可能保留模板的其他包内容。更新 NuGet 包也不会升级应用实际宿主的 .NET runtime；framework-dependent 应用应更新宿主，自包含应用和容器需重新发布或构建并部署。完整边界见 [安全说明](SECURITY.md#仍需由应用负责的边界)。
