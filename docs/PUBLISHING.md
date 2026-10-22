# 发布指南

发布入口是 [release.yml](../.github/workflows/release.yml)：向远端推送 `v*` tag 后，先执行 Windows 测试门禁，通过后在 Ubuntu 构建四包、推送 nuget.org，再创建 GitHub Release。发布历史见 [CHANGELOG](../CHANGELOG.md)，已核验的发布结果见 [评审台账](reviews/2026-09-07-review.md)。

## 首次配置（仓库或发布身份变更时复核）

- 在 nuget.org 为发布身份配置本仓库 `CSJ608/TemplateFrame` 的 Trusted Publisher，对应工作流文件 `release.yml`。正常发布由 `v*` tag 触发；补传符号包可从 `main` 手动运行同一 `release.yml`。两条路径均未指定 GitHub Environment；配置需与实际仓库和工作流身份匹配。
- 在 GitHub 仓库 Settings → Secrets and variables → Actions → **Variables** 配置 `NUGET_USER`，值为 nuget.org 用户名。工作流读取的是 `vars.NUGET_USER`，不是同名 Secret。
- 发布 job 使用 `id-token: write` 与 `contents: write`；手动补传的 `symbols` job 仅使用 `contents: read` 与 `id-token: write`，通过固定 SHA 的 `NuGet/login`（v1.2.0） 换取短时 API key。账号权限、Trusted Publisher 和变量值需发布者在平台上确认，不能仅凭仓库文件认定配置已完成。

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

使用 `global.json` 固定的 .NET SDK `10.0.402`；CI 另安装 `8.0.426` 提供已修补的 net8 运行时。使用系统级 `dotnet`，环境核验方式参见 [安全与开发环境](SECURITY.md)。在 Windows 执行下面的全目标验证（测试目标 `net8.0` / `net472` / `net10.0`；库资产为 `netstandard2.0` / `net462` / `net8.0`）。每步成功后再继续：

```powershell
dotnet restore TemplateFrame.slnx
./eng/audit-dependencies.ps1
dotnet format TemplateFrame.slnx --verify-no-changes --no-restore
dotnet build TemplateFrame.slnx -c Release --no-restore
dotnet test TemplateFrame.slnx -c Release --no-build

$PackageOutput = "artifacts/release-check-$ReleaseVersion"
dotnet pack src/TemplateFrame/TemplateFrame.csproj -c Release --no-build -o $PackageOutput
dotnet pack src/TemplateFrame.Word/TemplateFrame.Word.csproj -c Release --no-build -o $PackageOutput
dotnet pack src/TemplateFrame.Excel/TemplateFrame.Excel.csproj -c Release --no-build -o $PackageOutput
dotnet pack src/TemplateFrame.Excel.Simple/TemplateFrame.Excel.Simple.csproj -c Release --no-build -o $PackageOutput
```

Linux 本地测试分别使用 `dotnet test TemplateFrame.slnx -c Release --no-build -f net8.0` 和 `-f net10.0`，不能代替发布工作流的 Windows 全目标门禁。

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
| `test`（仅 tag push） | Windows | Version/tag 校验 → restore → 独立依赖审计 → Release build → 全目标测试 |
| `release`（仅 tag push，`needs: test`） | Ubuntu | restore → 独立依赖审计 → build → 四包 pack → 检查 NUGET_USER → OIDC 登录 → V3 推送四个 nupkg（`--no-symbols`）→ 单独推送四个 snupkg → 提取 CHANGELOG 段 → 创建 GitHub Release，附 nupkg/snupkg |
| `symbols`（仅 `main` 手动触发） | Ubuntu | 下载已有正式 Release 的四个 snupkg → 验证 GitHub 摘要及包内身份/PDB → 检查 NUGET_USER → OIDC 登录 → V3 仅推送四个 snupkg |

正常发布的主包和符号包均使用 NuGet V3 源 `https://api.nuget.org/v3/index.json`，主包步骤显式 `--no-symbols`，下一步单独推送 `.snupkg`，日志应分别出现四次主包和四次符号包上传。V2 源不支持 `.snupkg` 发布，GitHub Release 附有符号包也不能证明 NuGet 符号服务收到它们。两步均使用同一短期 key 和 `--skip-duplicate`；跳过已存在包不代表它与本次构建内容相同。发布后核对 Actions 的提交与结果、GitHub Release 八个资产、nuget.org 四包的实际版本和元数据，以及符号包上传结果和后续验证状态；只看到 tag 或工作流启动不等于发布完成。

## 补传已有正式 Release 的符号包

当主包已发布且 GitHub Release 已保留原始四个 `.snupkg`，可用 `release.yml` 的手动入口补传。先确认该工作流修复已经合并到 `main`，再在 Actions 的 Run workflow 中选择 **main**，输入已有正式 tag（例如 `v3.0.0`）。也可执行：

```powershell
gh workflow run release.yml --ref main -f release_tag=v3.0.0
gh run list --workflow release.yml --event workflow_dispatch
```

`release_tag` 仅接受 `vX.Y.Z`，不接受预发布或任意分支名。`symbols` job 仅允许 `refs/heads/main`；手动运行不会进入正常发布的 `test` / `release` jobs。此入口继续使用同一仓库、同一 `release.yml` 的 Trusted Publisher 身份与 `NUGET_USER`，安装固定 SDK 10.0.402 仅用于 NuGet 推送。

工作流通过当前仓库的 GitHub API 查询该 tag 的既有 Release，拒绝 draft / prerelease。四个资产必须各有一个精确名称：`TemplateFrame.<版本>.snupkg`、`TemplateFrame.Word.<版本>.snupkg`、`TemplateFrame.Excel.<版本>.snupkg`、`TemplateFrame.Excel.Simple.<版本>.snupkg`。下载时逐一核对 GitHub 提供的 SHA256 digest 与大小，再只读检查 ZIP 中的 nuspec ID/版本及 `netstandard2.0 / net462 / net8.0` 三组 Portable PDB；缺失摘要或任一检查失败都会在申请 NuGet key 前终止。

补传使用 Release 中已保存的原始符号包，不重建、不重新推送 nupkg、不创建 Release，也不修改 tag。流程通过 V3 源只推送校验完成的四个 snupkg，并允许跳过已存在的符号包。执行后检查四个文件的推送日志与 NuGet 后续符号验证状态；异步验证尚未结束时不要把上传成功写成索引已可用。

参考：[NuGet 符号包发布要求](https://learn.microsoft.com/en-us/nuget/create-packages/symbol-packages-snupkg)、[GitHub Release 资产 API](https://docs.github.com/en/rest/releases/assets)。
