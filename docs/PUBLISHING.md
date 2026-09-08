# 发布指南

发布入口是 [release.yml](../.github/workflows/release.yml)：向远端推送 `v*` tag 后，先执行 Windows 测试门禁，通过后在 Ubuntu 构建四包、推送 nuget.org，再创建 GitHub Release。发布历史见 [CHANGELOG](../CHANGELOG.md)，已核验的发布结果见 [评审台账](reviews/2026-09-07-review.md)。

## 首次配置（仓库或发布身份变更时复核）

- 在 nuget.org 为发布身份配置本仓库 `CSJ608/TemplateFrame` 的 Trusted Publisher，对应工作流文件 `release.yml`。当前工作流由 `v*` tag 触发，未指定 GitHub Environment；配置需与实际仓库和工作流身份匹配。
- 在 GitHub 仓库 Settings → Secrets and variables → Actions → **Variables** 配置 `NUGET_USER`，值为 nuget.org 用户名。工作流读取的是 `vars.NUGET_USER`，不是同名 Secret。
- 工作流声明 `id-token: write` 与 `contents: write`，通过 `NuGet/login@v1` 换取短时 API key。账号权限、Trusted Publisher 和变量值需发布者在平台上确认，不能仅凭仓库文件认定配置已完成。

## 日常发布：版本与说明

在仓库根目录操作。以下命令使用 **PowerShell**；`$ReleaseVersion` 是本次待发布版本，读取已由发布准备步骤确定的共享配置，不在示例里写死版本号：

```powershell
[xml]$ReleaseProps = Get-Content -Raw src/Directory.Build.props
$ReleaseVersion = [string]$ReleaseProps.Project.PropertyGroup.Version
if ([string]::IsNullOrWhiteSpace($ReleaseVersion)) { throw '共享 Version 不能为空' }
$ReleaseTag = "v$ReleaseVersion"
$ReleaseTag
```

发布前逐项核对：

1. 四包 `TemplateFrame`、`TemplateFrame.Word`、`TemplateFrame.Excel`、`TemplateFrame.Excel.Simple` 的版本统一来自 [src/Directory.Build.props](../src/Directory.Build.props) 的 `<Version>`，各项目不另行覆盖。
2. 将本次 Unreleased 内容归档为 `## [版本号] - YYYY-MM-DD`，其中版本号等于 `$ReleaseVersion`；保留新的 `## [Unreleased]` 供后续修改记录。不要改动历史版本段。
3. 共享 `Version`、CHANGELOG 版本段和去掉 `v` 前缀的 tag 必须一致。工作流强制校验 **Version ↔ tag**；CHANGELOG 段缺失时会回退到自动生成 Release 正文，因此 CHANGELOG 一致性仍需人工检查。
4. 审阅最终差异，完成下面的本地验证。确认待发布提交、远端 CI 结果及 tag 尚未使用，再执行发布动作。

## 本地验证（四包）

使用支持 `.slnx` 的 SDK；工作流使用 .NET SDK `9.0.x`。在 Windows 执行下面的全目标验证（测试目标 `net8.0` / `net472`；库资产为 `netstandard2.0` / `net462` / `net8.0`）。每步成功后再继续：

```powershell
dotnet restore TemplateFrame.slnx
dotnet format TemplateFrame.slnx --verify-no-changes --no-restore
dotnet build TemplateFrame.slnx -c Release --no-restore
dotnet test TemplateFrame.slnx -c Release --no-build

$PackageOutput = "artifacts/release-check-$ReleaseVersion"
dotnet pack src/TemplateFrame/TemplateFrame.csproj -c Release --no-build -o $PackageOutput
dotnet pack src/TemplateFrame.Word/TemplateFrame.Word.csproj -c Release --no-build -o $PackageOutput
dotnet pack src/TemplateFrame.Excel/TemplateFrame.Excel.csproj -c Release --no-build -o $PackageOutput
dotnet pack src/TemplateFrame.Excel.Simple/TemplateFrame.Excel.Simple.csproj -c Release --no-build -o $PackageOutput
```

Linux 本地测试使用 `dotnet test TemplateFrame.slnx -c Release --no-build -f net8.0`，不能代替发布工作流的 Windows 全目标门禁。

打包后分别核对四包，而不是仅检查核心与 Word：

| 包 ID | README 来源 | 预期产物 |
|---|---|---|
| `TemplateFrame` | [根 README](../README.md) | `TemplateFrame.<版本>.nupkg` / `.snupkg` |
| `TemplateFrame.Word` | [Word README](../src/TemplateFrame.Word/README.md) | `TemplateFrame.Word.<版本>.nupkg` / `.snupkg` |
| `TemplateFrame.Excel` | [Excel README](../src/TemplateFrame.Excel/README.md) | `TemplateFrame.Excel.<版本>.nupkg` / `.snupkg` |
| `TemplateFrame.Excel.Simple` | [Simple README](../src/TemplateFrame.Excel.Simple/README.md) | `TemplateFrame.Excel.Simple.<版本>.nupkg` / `.snupkg` |

- 本次版本应有 4 个 nupkg 和 4 个 snupkg；检查输出目录，避免把其他版本旧包当成本次结果。
- 按 ZIP 查看每个包：nuspec 的 ID/版本/依赖组正确，普通包包含三个目标的 DLL、XML 文档、en 卫星程序集，以及对应 README 和 icon；符号包包含三个目标的 PDB。
- README 与工作区来源一致；没有测试依赖、测试文件、日志或临时验证文件混入。只有最终提交的发布工作流产物才作为正式制品，本地包用于验证。

## 发布动作（单独执行）

最终内容完成审阅、提交并推送，且该提交的 CI 成功后，在同一 PowerShell 会话重新核对 `$ReleaseVersion` / `$ReleaseTag` 与提交内容。下面的 tag 推送会触发真实发布：

```powershell
git tag -a $ReleaseTag -m "Release $ReleaseVersion"
git push origin $ReleaseTag
gh run list --workflow release.yml
```

## 工作流与发布后核验

| Job | 环境 | 实际顺序 |
|---|---|---|
| `test` | Windows | Version/tag 校验 → restore → Release build → 全目标测试 |
| `release`（`needs: test`） | Ubuntu | build → 四包 pack → 检查 NUGET_USER → OIDC 登录 → NuGet push → 提取 CHANGELOG 段 → 创建 GitHub Release，附 nupkg/snupkg |

NuGet push 使用 `--skip-duplicate` 跳过已存在包，不代表已存在包与本次构建内容相同。发布后核对 Actions 的提交与结果、GitHub Release 的四包及四符号包，以及 nuget.org 四包的实际版本和元数据；只看到 tag 或工作流启动不等于发布完成。
