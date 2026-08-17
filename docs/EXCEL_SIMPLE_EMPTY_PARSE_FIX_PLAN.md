# Excel.Simple 解析「有数据却为空」问题清单与修复计划

> 状态：**已实施（2026-08-17）**——评审通过后代码与回归测试已落地，14 个探针场景全部修复；待用户环境 `dotnet test` 验证（本会话沙箱无法启动 testhost）
> 范围：`src/TemplateFrame.Excel.Simple`（`SimpleExcel.Read` / `SimpleExcelContract.Read`）
> 依据：静态分析 + 探针实测（探针位于 `%TEMP%\dsh-81QTv0\tf-excelsimple-probe\`，未改动仓库）

---

## 一、问题清单（已实测复现）

| # | 问题 | 严重度 | 复现场景 | 根因（代码位置） |
|---|---|---|---|---|
| P1 | **命名区域 `TF_Table` 过窄/错位时静默丢数据** | 🔴 高 | S1：区域只盖表头、数据在下方 → 数据行=0；S2：区域指向错位 → 完全为空；契约路径 C1/C2 同样为空 | `SimpleExcel.Read` 拿到命名区域后完全按 `StartRow/EndRow` 读（L198-225），数据循环 `r = headerRow+1 .. endRow`（L242）；区域覆盖不到的数据**没有任何向下扫描兜底**。最典型真实场景：`BuildTemplate()` 只写表头（`TF_Table=A1:C1`），用户手工在 Excel 里填数后回读 → 0 行 |
| P2 | **共享字符串表头被读成索引号** | 🔴 高 | S5：Excel 生成文件（字符串默认共享字符串）+ 无命名区域 → 表头变成 "0\|1\|2"；契约回退 C3 → 每行是空字典，列全部匹配失败 | `GetCellText` 对 SharedString 单元格直接返回 `CellValue.Text`（索引），从不解析（L401-407）。Excel/WPS 生成的文件几乎必中 |
| P3 | **富文本共享字符串单元格读成 null** | 🟠 中 | S4：整行富文本 → 行被"全空行"逻辑丢弃 → 结果为空；S8：行内混有 → 该列静默变 null | `ReadCellValue` 取 `SharedStringItem.Text`（直接 `<t>` 子元素），富文本项只有 `<r>` 片段 → null（L338-339）；再被全空行跳过（L251-254） |
| P4 | **行定位强依赖 RowIndex，缺失即失败** | 🟠 中 | S3：命名区域正确但行缺 `r` 属性 → 完全为空；S7：回退路径 → 直接抛 NullReferenceException | `FindCell` 按 `RowIndex` 精确匹配（L409-419）；回退路径 `RowIndex!.Value` 无空值保护（L221/224）。OOXML 规范中 `r` 属性可选 |
| P5 | **回退"首非空行"假设脆弱（标题行/装饰行在上方）** | 🟡 低 | S6：第 1 行是"物料清单"标题 → 被当表头，`colCount` 取该行单元格数 → 数据被截断成 1 列 | 回退逻辑 L214-224：首非空行即表头，无"哪个更像表头"的判断 |

对照组 S0/C0（框架自写自读）与现有全部单元测试语义正常，说明问题集中在**外部/手工编辑文件**的读入路径。

## 二、修复计划

### 设计原则
- **不改变正确文件的现有行为**：框架 `Write` 产物的区域天然精确，修复后结果与今天一致（现有测试保持全绿）。
- 修复只作用于"区域读不到数据 / 数据在区域外 / 非标准文件"这些异常形态。
- 契约路径与裸 `SimpleExcel.Read` 共用同一套定位逻辑，避免两处行为分叉。

### P1（核心修复）——命名区域兜底扩展
`SimpleExcel.Read` 命名区域分支改造（`SimpleExcel.cs`）：
1. 区域表头行为空（或区域 0 数据行）→ **回退到现有"首非空行"扫描逻辑**（修复 S2 错位场景）。
2. 区域读到 0 行数据但区域下方（同列跨度内）仍有非空行 → **把 `endRow` 顺延到最后一个非空行再读**（修复 S1）。
3. **P1-3（评审采纳：启用，不加护栏）**：区域尾行下方紧邻仍是非空数据行时也顺延（修复"区域只盖住部分数据"的变体）。评审结论：现有回退路径本就"读到工作表最后一行"（`SimpleExcel.cs:224`），"下方第二表格被并入"是回退路径**今天已有**的行为，P1-3 只是让命名区域分支与之一致；本包哲学是单表简单场景，一致性比"精确截断"更重要，不加"连续 N 个全空行即停"护栏，保持规则简单可预测。
- 契约路径 `ReadByLayout`（`SimpleExcelContract.cs` L306-349）同样应用"0 行则向下顺延"规则，或在无列定义名时自然落入已修复的 `SimpleExcel.Read` 文本回退（C2 自动受益）。

### P2 + P3（同一根因）——共享字符串解析统一
新增内部助手 `ResolveSharedString(WorkbookPart, int index)`：
- 优先取直接 `<t>` 文本；无直接 `<t>` 时**拼接所有 `<r>` 片段文本**（富文本，修复 P3）；
- `GetCellText` 与 `ReadCellValue` 共用该助手（修复 P2 表头索引号问题）；
- **R2（评审采纳）**：`GetCellText(Cell?)` 签名变更——解析共享字符串需要 `WorkbookPart`，改为 `GetCellText(WorkbookPart, Cell?)`（internal 方法，无公开 API 破坏），调用点同步更新：表头检测（L215）、表头读取（L238）、契约 `ResolveColumnLayout`（L299）。
- **R1（评审采纳，性能）**：共享字符串表**一次性物化**。`Elements<SharedStringItem>().ElementAtOrDefault(index)` 是 O(n) 查找，修复后每个字符串单元格都走这条路，大文件导入退化为 O(行×列×SST 大小)。在 `Read()` 入口把 `SharedStringTable` 物化成 `IReadOnlyList<SharedStringItem>` 一次性传下（`ReadCellValue` / `GetCellText` / `ResolveSharedString` 改为接收物化列表或 `WorkbookPart` 携带）。
- 附带收益：P2 修复后，契约文本匹配对 Excel 生成文件（S5/C3）恢复有效。

### P4——RowIndex 缺失兜底（防御性）
- `FindCell`：行无 `RowIndex` 时按 `CellReference` 列字母 + 文档顺序兜底匹配；
- 回退路径 `RowIndex!.Value` 加 `?? 文档序号` 空值保护，消除 NRE（S7）；
- **R3（评审采纳）**：`SimpleExcelContract.Validate` 异常捕获列表（`SimpleExcelContract.cs:166-170`）不含 NRE，P4 根治后此隐患自然消除；测试计划补充"Validate 对 S7 形态文件不再抛异常"断言。

### P5——回退表头选择启发式（评审采纳：改窄后纳入，见第六节决策）
- **改窄版**（不做全局 argmax，评审指出 argmax 会把"用户多填几列的数据行"误判为表头）：**跳过"仅有 1 个非空单元格"的前导行**（标题行/装饰行特征），以其后首个多单元格行为表头；
- 属启发式，误判风险低但存在，若实施中遇阻可顺延到单独迭代。

### 测试计划（新增回归测试，`test/TemplateFrame.Excel.Simple.Tests`）
1. `Read_NamedRangeTooNarrow_StillReadsDataBelow`（S1）
2. `Read_NamedRangeMisplaced_FallsBackToFirstNonEmptyRow`（S2）
3. `Read_RowsWithoutRowIndex_StillReads`（S3，含回退路径不再抛 NRE）
4. `Read_RichTextSharedStrings_ResolvesText`（S4/S8）
5. `Read_ExcelSharedStringHeaders_ResolvesRealText`（S5）
6. `ContractRead_NarrowNamedRange_ReturnsDataRows`（C1/C2）
7. `ContractValidate_RowsWithoutRowIndex_DoesNotThrow`（R3，S7 形态）
8. `Read_LeadingSingleCellTitleRow_SkippedAsHeader`（P5 改窄版，S6）
9. 现有测试全部保持通过（S0/C0 行为不变）

### 涉及文件
- `src/TemplateFrame.Excel.Simple/SimpleExcel.cs`（主要改动：命名区域兜底、共享字符串物化与解析、RowIndex 兜底、表头启发式）
- `src/TemplateFrame.Excel.Simple/SimpleExcelContract.cs`（`ReadByLayout`、`ResolveColumnLayout` 调用点随签名变更）
- `test/TemplateFrame.Excel.Simple.Tests/`（新增回归测试）
- **`CHANGELOG.md` 必写（R4，评审采纳）**：区域自动顺延是用户可感知的读取容错行为变化，需文档说明；插件 README 同步

---

## 三、待审核决策点（2026-08-17 已由评审给出意见，见 4.4 / 六）

1. **P1-3 是否启用**：评审建议**启用**（不加护栏）→ 已采纳。
2. **P5 是否纳入本次**：评审建议**改窄后纳入**（跳过仅 1 个非空单元格的前导行，不做 argmax）；时间紧可顺延 → 计划按"改窄版纳入"写，最终范围待用户确认（见第六节）。

---

## 四、评审意见（2026-08-17）

### 4.1 核查结论：**分析全部属实**

评审人对源码逐行核对（文档引用行号全部准确）、独立复跑探针（S0–S8 / C0–C3 共 12 个场景，结果与本文件"一、问题清单"逐条吻合）、并运行现有测试（`TemplateFrame.Excel.Simple.Tests` 30/30 通过，S0/C0 对照组正常）。

- P1 的典型场景已源头确认：`SimpleExcelTemplateService.BuildTemplate()`（`SimpleExcelTemplateService.cs:48-54`）以空 `FillData` 调 `Write`，`endRow = rowIndex - 1`（`SimpleExcel.cs:142`）落在表头行 → `TF_Table = $A$1:$C$1`；用户在 Excel 填数不会自动扩展命名区域，回读必得 0 行。**这是本插件核心工作流（生成模板 → 用户填写 → Parse 回读）的断裂**，🔴 高严重度成立，甚至可算产品级缺陷。
- P4 补充依据：ECMA-376 中行/单元格 `r` 属性可选，缺失时按"前一个 +1"推断是**规范行为**而非启发式，修复方案方向正确。

### 4.2 方案总体评价：正确、分层清晰、可实施

1. **根因归并准确**：P2+P3 识别为同一根因，统一 `ResolveSharedString` + 两处共用是教科书式修法；P2 修好后 C3 契约文本匹配自动恢复的连带收益判断也对。
2. **约束意识好**："不改变正确文件现有行为" + S0/C0 回归 + 决策点诚实标注权衡。
3. 契约路径与裸 `SimpleExcel.Read` 共用定位逻辑，避免行为分叉。
4. 测试计划与复现场景一一对应，可直接落地。

### 4.3 实施前建议补充/修正

| # | 建议 | 归属 |
|---|---|---|
| R1 | **共享字符串物化（性能）**：`Elements<SharedStringItem>().ElementAtOrDefault(index)` 是 O(n) 查找，修复后每个字符串单元格都走这条路，大文件导入退化为 O(行×列×SST大小)。建议在 `Read()` 入口把共享字符串表物化成 `IReadOnlyList<SharedStringItem>` 一次性传下。既然动的就是这段代码，顺手做掉 | P2+P3 |
| R2 | **点明 `GetCellText` 签名变更**：解析共享字符串需要 `WorkbookPart`，现有 `GetCellText(Cell?)` 必须加参数（internal 方法，无公开 API 破坏），计划中写明以免实施时犹豫 | P2 |
| R3 | **Validate 路径补一条断言**：其异常捕获列表（`SimpleExcelContract.cs:166-170`）不含 NRE，P4 根治后自然失效；测试计划建议加"Validate 对 S7 形态文件不再抛异常" | P4 / 测试计划 |
| R4 | **CHANGELOG 必写**（非"视情况"）：区域自动顺延是用户可感知的读取容错行为变化，需文档说明 | 收尾 |

### 4.4 决策点意见

1. **P1-3：建议启用**。现有回退路径本就"读到工作表最后一行"（`SimpleExcel.cs:224`），"下方第二表格被并入"是回退路径**今天已有**的行为，P1-3 只是让命名区域分支与之一致；本包哲学是单表简单场景，一致性比"精确截断"更重要。若仍担心，可加"连续 N 个全空行即停"护栏，但建议保持规则简单可预测，不加。
2. **P5：建议改窄后再纳入（或顺延）**。"选非空单元格数最大的首行"的 argmax 规则**误判风险并不低**——数据行常因用户多填几列而比表头行更宽，会把第一数据行选成表头。更稳的最小修法：**跳过"仅有 1 个非空单元格"的前导行**（标题行特征），不做全局 argmax。若时间紧，顺延到单独迭代也可接受。

**评审结论：按本计划实施，采纳 R1–R4 与上述两个决策点意见。**

---

## 五、评审采纳记录（2026-08-17）

| 项 | 评审意见 | 采纳结果 |
|---|---|---|
| R1 | 共享字符串表一次性物化（性能） | ✅ 已并入 P2+P3 节：`Read()` 入口物化为 `IReadOnlyList<SharedStringItem>` 传下 |
| R2 | 点明 `GetCellText` 签名变更（加 `WorkbookPart` 参数） | ✅ 已写入 P2+P3 节 |
| R3 | Validate 补"S7 形态不再抛异常"断言 | ✅ 已写入 P4 节与测试计划 #7 |
| R4 | CHANGELOG 必写（非"视情况"） | ✅ 已写入涉及文件节 |
| P1-3 | 启用，不加"连续全空行即停"护栏 | ✅ 已写入 P1 节 |
| P5 | 改窄：跳过仅 1 个非空单元格的前导行，不做 argmax | ✅ 已写入 P5 节（范围待用户确认，见第六节） |

## 六、实施前待用户确认

1. **P5 范围**：按评审"改窄版"纳入本次实施，还是顺延到单独迭代？
2. **实施启动**：确认后按本计划（P1–P5 + R1–R4）开始修改代码并补回归测试。

---

## 七、实施记录（2026-08-17）

**已确认**：P5 改窄版纳入本次；立即实施。

**代码落地**：
- `src/TemplateFrame.Excel.Simple/SimpleExcel.cs`：`Read` 重构（命名区域表头为空回退 + endRow 顺延）；`GetCellText` / `ReadCellValue` 支持共享字符串物化解析（R1）与富文本拼接（P3）；`GetRowIndex` / `BuildRowLookup` / `FindCell` 行索引推断（P4）；`FindHeaderRowIndex` 跳过单格前导行（P5）
- `src/TemplateFrame.Excel.Simple/SimpleExcelContract.cs`：`ResolveColumnLayout`（共享字符串表头 + 区域表头为空回退）、`ReadByLayout`（endRow 顺延 + 物化共享字符串）
- `test/TemplateFrame.Excel.Simple.Tests/ReadToleranceTests.cs`：新增 11 个回归测试（S1–S8 / C1–C3 / R3）
- `CHANGELOG.md`（R4 必写）与插件 `README.md` 同步

**验证**（沙箱内无法启动 testhost，改用临时工程编译改动源码 + 探针复跑全部场景）：
- 改动源码编译通过（临时类库，0 警告 0 错误）；回归测试文件编译通过
- 探针 14 场景（S0–S10 / C0–C3）全部修复：S1 3 行、S2 表头正确、S3 2 行、S4 富文本拼接、S5 真实表头、S6 标题行跳过、S7 不再 NRE、S8 富文本列有值、S9 类型化往返、S10 null 补齐；C1/C2 3 行、C3 值完整；对照组 S0/C0 行为不变
- 遗留：用户环境跑 `dotnet test` 全量回归（现有 30 用例 + 新增 11 用例）
