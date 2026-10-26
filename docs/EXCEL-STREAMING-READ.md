# Excel 流式导入实施记录（2026-10-26）

## 用户授权与范围

- 优先解决 Excel 导入的大内存问题；导出逐行 CPU 优化延后，不做模板支持范围扩大及 Office/WPS 真实模板验收。
- 用户已明确允许可配置目录的临时磁盘文件，用于共享字符串正文/索引及不可 Seek 输入暂存。
- 旧 Read/Parse 接口、语义及默认数据限额不变。新入口显式 opt-in，不将返回完整集合伪装成流式读取。
- 无全局并发信号量；调度由宿主负责。按小批实现、验证、本地 Git 提交，不推送。
- 延续中等推理；不新增不必要的代码注释，清晰命名、小函数、控制参数数量。
- 用户要求接近上下文占用 70% 时新建独立会话，不允许 fork。工具未提供可核验的当前占用比例，不得编造百分比；独立接续需完整交接，旧会话停止代码写入。

## 设计方向

先实现 `SimpleExcel.ReadStreaming`，逐行回调或单次可释放枚举消费，不累积返回行。普通读取跑通并验证后，再复用底层读取能力接入契约及模板读取；不要把后两者列为已经完成。

工作表和共享字符串通过流式 XML 读取，避免访问 Worksheet/SharedStringTable 的完整 DOM。共享字符串索引本身也落盘（每项 8 字节偏移 + 4 字节 UTF-8 长度），正文为 UTF-8，内存仅保存受字符数及项数双重约束的 LRU 缓存。元数据另设有限 XML 字符预算，防止为了百万行提高工作表限额时也无限放大样式/工作簿 DOM。

可 Seek 源优先直接读；非 Seek 源按输入限额写临时文件。临时文件累计字节预算覆盖输入暂存、字符串正文与索引。随机文件名、CreateNew、FileShare.None、DeleteOnClose；不扫描/清理调用方其他文件。目录由宿主配置及保护，文件可能含业务数据，不是加密存储或安全擦除。正常释放/异常/取消应清理；进程强杀、系统崩溃或文件系统故障不能承诺跨平台零残留。

后续必须明确：表头回退可能需要额外一遍流式扫描；旧读取接受乱序/重复行并排序，新入口不能悄悄积累全部行来兼容，需显式拒绝或设计磁盘处理；单元格列乱序可以在当前行内处理。公式只读取缓存值，不新增重算；日期/字符串/空值语义应与旧普通读取对照，契约/模板原本不同的转换规则不能混同。错误/取消前回调可能已处理部分数据，事务或回滚由调用方负责。警告也不能无限积累。

## 第一批：临时存储与磁盘字符串索引

新增（尚未接入公开读取入口）：

- `SimpleExcelStreamingReadOptions`：原 ResourceLimits；临时目录；默认 256 MiB 临时预算；默认缓存 1 Mi 字符/1024 项；默认元数据 1 Mi 字符限额。
- `SimpleExcelTemporaryStorage`：创建仅本次拥有的临时文件、累计写入预算、非 Seek 输入暂存、释放时尝试关闭全部文件。
- `SimpleExcelDiskStrings`：正文和固定长度索引均落盘，随机读取，有界 LRU；32,767 字符单项上限，不缓存过大的字符串。
- `StreamingStorageTests`：随机逆序读取、索引/UTF-8 预算、缓存边界、空串、非法索引、超长字符串、输入限额/磁盘限额/取消清理及调用方源所有权。

第一批验证：Release 全量构建零警告零错误；net472/net8/net10 全量共 2,030 次测试通过（Core399、Word492、Excel702、Simple437），含新增8场景×3框架。已执行方案格式化。尚未实现公开 ReadStreaming 入口、ZIP/XML 逐行读取、契约/模板接入或导入容量实验；这些是下一批，不得将底层存储完成视为流式导入完成。

## 当前接续点（普通读取已完成，契约入口已接入，模板待做）

2026-10-26 接续批次 `7114a1e` 新增 `SimpleExcelContract.ReadStreaming` 与 `SimpleExcelContractReadRow`，复用直接 ZIP/XML 和磁盘 SST 读取，不构造完整表格。抽取定义名布局映射核心供旧接口和新接口共用，保留重复定义名、位置冲突、无效列范围回退等旧语义；定义名路径缺键补 null/契约列全空跳过，表头回退路径仅匹配键且保留空字典行。新增 12 场景覆盖旧接口逐键对照与非 Seek 成功、提前停止、取消、消费者异常、磁盘预算清理。模板接口仍待完成，不代表整体交付。

本批全方案 net472/net8/net10 共 2,144 次测试通过（Core399、Word492、Excel702、Simple551）；方案已 format，独立探针构建零警告零错误并 whitespace format，git diff --check 通过。探针新增 `contract-streaming-read`，逐键逐格验证。契约千/万/十万/百万行 inline/shared 均已按序验证，.NET 10 百万 inline 30.152 秒/60.32 MiB，shared 108.015 秒/57.89 MiB，均零临时文件残留，详见[契约容量证据](reviews/contract-streaming-read-2026-10-26/README.md)。

`75d615c` 将四个底层类提取到 `src/Shared/Excel`，命名空间 `TemplateFrame.Internal.Excel`，重命名为 `ExcelReadPackage`、`ExcelBoundedXml`、`ExcelTemporaryStorage`、`ExcelDiskStrings`。两 Excel csproj 以链接源码编译，不增加 NuGet 包或项目间依赖。内部配置接口 `IExcelStreamingReadSettings` 由现有公开 Simple 选项实现；模板选项后续可实现相同接口。提取后全量 2,144 测试再通过，方案及探针 Release 构建零警告零错误。逐行 worksheet/SST 的扫描逻辑仍在 `SimpleExcelStreamingReader`，尚未提取或接入模板。

契约补充验证：.NET 8 百万 inline 51.626 秒/56.83 MiB，shared 102.341 秒/56.81 MiB；8 个独立千行 shared 小任务并发逐格通过，批次 0.838 秒/66.58 MiB，全部零残留。首轮 .NET 8 shared 启动前内存保护拒绝，未启动子进程；原目录保留，系统空闲恢复后新 r2 目录成功，未降低任何保护或关闭用户应用。详情与原始归档见契约容量证据。

普通流式读取百万行已验证。详见 [容量证据](reviews/streaming-read-2026-10-26/README.md)：六列逐格校验；.NET 10 inline 32.94 秒/60.22 MiB 峰值工作集，4,000,006 项高基数 shared-string 108.50 秒/57.82 MiB；.NET 8 对应 32.71 秒/58.23 MiB、126.90 秒/56.23 MiB。SST 临时内容 136,000,084 字节，所有成功样例退出后零文件。万行到百万行的内存没有线性增长。保护参数未降低；首次启动前因可用内存不足被拒绝，正常关闭构建服务器后新 case 重试。生成样本的内存不能混作读取结果。

普通接口阶段公开同步回调入口、临时写入计费结果、中英文使用文档已实现。补充 XML 底层每次读入的取消/条目字节检查、损坏 ZIP/XML、预取消不消费源和无效临时目录测试。当时全方案三框架 2,108 测试通过；当前计数见上。容量探针不在方案内，需独立构建/format。

下一批务必继续，不能把普通读取当成全部交付：

1. 契约入口、旧语义对照和容量验证已有独立证据，不必重做；后续共享底层修改仍需回归契约与普通入口。
2. 接入模板流式解析。复用 `src/Shared/Excel`，继续提取真正共用的 worksheet/SST 扫描逻辑；不让 Excel 项目依赖 Simple 项目。阅读 `ExcelTemplateParser` 及相关解析操作/布局类，保留类型转换、缺键和诊断语义；警告采用回调或明确有界输出。旧入口及默认限额不变。
3. 保持真正 ZIP/XML 流式和有界内存，不能通过 OpenXml Packaging 再物化工作表或 SST；拒绝倒序/重复行的行为已有文档。常规 transitional OOXML 样本已测；新 reader 对非预期 worksheet/SST 根命名空间显式报错，不可悄悄当空表。
4. 补充契约/模板各自对照、生命周期/异常与容量验证，必要时安全小任务并发；不要执行百个百万行并发或 Word 导入。原保护/规模上限不变。
5. 每完成并验证小批本地提交、不推送。后续按用户约定仅新建独立会话接续，不 fork；工具没有可核验上下文百分比，不编造 70% 读数。

模板旧语义排查（仅分析，尚未实现）：

- 定义名 `TF_Key`、`TF_TableKey_ColumnKey`，首匹配；`ExcelNamedRangeLocator.ParseReference` 与 Simple 地址解析规则不同，不能直接替换。ResolveWorksheetPart 按工作表名大小写精确匹配，不存在不回退首表。
- 表按已匹配列的最大范围长度输出，空逻辑行也保留；每列从自己的 Start.Row 加逻辑偏移读取，即使超出该列自己的 End.Row 仍读。所有列使用最后匹配列的工作表（旧行为）。可按起始行偏移分组建前向游标，不能为错位列缓存整个表；需要保留多表与不同起点的对照测试。
- 旧 CellLookup 只索引显式 RowIndex，精确 CellReference 首值；普通读取对缺省行号的推断不应自动套用。新接口拒绝重复/倒序行可延续已约定边界，须文档说明。
- 布尔值不经 ValueType 转换；只有显式 t=n 且契约类型 DateTime 才按 OADate 转换，不按样式。其他文本交给 ContractTextParser，保留失败时原字符串和 ConversionFailed 各诊断字段。SST 旧模板只读 item.Text?.Text，富文本无直属 Text 时为 null；Simple 则拼接富文本。共享磁盘字符串目前只支持非 null string，不能把模板 null 偷换成空串。inline 也只读直属 Text，再回退 CellValue/空串。
- 标量无命名单元格/无实际 cell 不返回键；图片复用现有 oneCellAnchor、精确零基位置、首 drawing/anchor、embed 关系与 MaxImageBytes 规则，不扩展 twoCellAnchor 等模板能力。不得通过 SpreadsheetDocument 打开全包来读图片；共享包目前只公开 XML/元数据，可能需增加有预算的二进制部件读取。
- ContractTextParser 当前诊断接收 List；可用逐单元格清空的小列表再回调或适当提取诊断投递接口，禁止全表警告累积。错误/取消/消费者异常必须清理，保持部分业务回调的责任说明。

以下为历史批次和原始验收清单，状态以本节为准。

第三批已接入 `SimpleExcel.ReadStreaming` 同步回调入口：返回只含表头/已消费行数/是否完成的结果；回调返回 false 可提前停止。当前行及单项共享字符串 XML 受新增 `MaxRowCharacters`（默认 1 Mi 字符）限制，避免单行 DOM 无界增长；仅对单行构造 OpenXml Row，不打开 SpreadsheetDocument，不构造完整 Worksheet/SST。共享字符串正文/索引落盘；只保留当前行、表头、有界字符串缓存、受元数据限额约束的样式缓存。乱序或重复物理行号显式拒绝；列乱序和重复单元格首值保持旧语义。普通读取的数字/日期/布尔/富文本/公式缓存/定义范围后继续读取/表头回退已与旧 Read 对照；非 Seek、提前停止、消费者异常、取消、磁盘和单行预算、空值物理行计费测试已通过。全方案三框架 2,096 次测试通过；已执行 dotnet format。契约和模板入口、容量验证仍待完成，不应把本批当作全部交付。

第二批底层包读取已实现 `SimpleExcelReadPackage`：直接 ZIP 逐缓冲解压校验全部条目，保留压缩输入/单条目/累计解压/条目数预算与取消检查；XML 禁用 DTD，元数据使用独立限额；解析部件相对、绝对和转义关系 URI，忽略外部关系，拒绝重复 ZIP 部件。没有使用 OpenXml Packaging 或整工作表缓冲。9 场景 × net472/net8/net10 全部通过。此批仍不是完整导入入口。

1. 接入流式 ZIP/XML 读取，保持包大小/解压实际字节预算及 XML 安全设置；不使用整部件 MemoryStream 作为隐性缓冲。
2. 旧 Read 对照：表头定位/回退、稀疏/空行、数字/日期/布尔、共享字符串/富文本、公式缓存、首个重复单元格、不可 Seek 输入、损坏包和取消/消费方异常清理。
3. 契约与模板接口的列映射、定位和转换诊断复用；警告通过回调或有界收集，避免新内存增长源。
4. 扩展独立 CapacityProbe。必须同时覆盖 inline string 和高基数 shared string，逐格校验，先千/万/十万，再百万行；不能只测本库 inline-string 导出文件。
5. heap5GiB/private6GiB/系统空闲最低4GiB/180秒保护不变，新 case 目录，验证磁盘峰值及最终清理。Excel最多百万行；100并发只按安全小任务验证，不等于100个百万行任务。
6. 全量三框架构建/测试/format，归档证据及使用/失败/存储责任文档。

导出工作已完成，起点 `a845ed9`，详情见 `docs/EXCEL-STREAMING.md` 和 `docs/reviews/template-streaming-2026-10-25/`，不要重复导出优化或向旧会话反馈。
