using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel;

/// <summary>
/// Excel 模板校验（迭代 8 兜底）：枚举 <c>TF_</c> 命名区域，按契约报告 Missing / WrongType / Ambiguous，
/// Extra 只告警放行（见设计文档 §5.3）。
/// </summary>
public sealed class ExcelTemplateValidator
{
    /// <summary>校验 .xlsx 模板与契约是否匹配。</summary>
    public TemplateValidationResult Validate(Stream template, TemplateContract contract)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(contract);
        if (template.CanSeek)
        {
            template.Position = 0;
        }

        try
        {
            using var document = SpreadsheetDocument.Open(template, false);
            if (document.WorkbookPart is null)
            {
                return Invalid("文档缺少工作簿部分（xl/workbook.xml）。");
            }

            return ValidateCore(document.WorkbookPart, contract);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or InvalidOperationException or FileFormatException)
        {
            return Invalid("无法打开 .xlsx（不是有效的 OOXML 包）：" + ex.Message);
        }
    }

    private static TemplateValidationResult ValidateCore(WorkbookPart workbookPart, TemplateContract contract)
    {
        var issues = new List<TemplateValidationIssue>();
        var names = ExcelNamedRangeLocator.FindAll(workbookPart);
        var knownNames = new HashSet<string>(StringComparer.Ordinal);

        // 1) 契约内部 Key 唯一性（含表格列 Key）
        foreach (var duplicate in contract.EnumerateTagKeys()
                     .GroupBy(k => k)
                     .Where(g => g.Count() > 1))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Invalid,
                Key = duplicate.Key,
                Message = "契约内部 Key 重复：" + duplicate.Key + "（命名区域必须全局唯一）。",
            });
        }

        // 2) 命名区域名全局唯一（workbook.xml definedName 本身要求唯一，兜底检测）
        foreach (var group in names.GroupBy(n => n.Name).Where(g => g.Count() > 1))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Ambiguous,
                Key = group.Key,
                Message = "命名区域 " + group.Key + " 重复出现。",
            });
        }

        // 3) 逐元素校验
        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TextElement text:
                    knownNames.Add(ExcelNamedRangeLocator.ElementName(text.Key));
                    CheckTextElement(text, names, issues);
                    break;

                case ImageElement image:
                    knownNames.Add(ExcelNamedRangeLocator.ElementName(image.Key));
                    CheckImageElement(image, workbookPart, names, issues);
                    break;

                case TableElement table:
                    foreach (var column in table.Columns)
                    {
                        knownNames.Add(ExcelNamedRangeLocator.TableColumnName(table.Key, column.Key));
                    }

                    CheckTableElement(table, workbookPart, names, issues);
                    break;
            }
        }

        // 4) 契约外命名区域：默认放行（告警）
        foreach (var match in names)
        {
            if (!knownNames.Contains(match.Name))
            {
                issues.Add(new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Extra,
                    Key = match.Name,
                    Message = "模板含契约外命名区域：" + match.Name,
                    Severity = TemplateValidationSeverity.Warning,
                });
            }
        }

        return new TemplateValidationResult { Issues = issues };
    }

    private static void CheckTextElement(
        TextElement element,
        IReadOnlyList<NamedRangeMatch> names,
        List<TemplateValidationIssue> issues)
    {
        var name = ExcelNamedRangeLocator.ElementName(element.Key);
        if (names.All(n => n.Name != name))
        {
            // 可选元素缺失只告警（模板仍有效），必填缺失才失败
            issues.Add(Missing(
                element.Key,
                "缺少文本元素 " + element.Key + "（" + element.DisplayName + "）对应的命名区域。",
                element.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning));
        }
    }

    private static void CheckImageElement(
        ImageElement element,
        WorkbookPart workbookPart,
        IReadOnlyList<NamedRangeMatch> names,
        List<TemplateValidationIssue> issues)
    {
        var name = ExcelNamedRangeLocator.ElementName(element.Key);
        var match = names.FirstOrDefault(n => n.Name == name);
        if (match is null)
        {
            issues.Add(Missing(
                element.Key,
                "缺少图片元素 " + element.Key + "（" + element.DisplayName + "）对应的占位图锚定格。",
                element.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning));
            return;
        }

        var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
        var worksheetPart = ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart is null)
        {
            issues.Add(WrongType(element.Key, "图片元素 " + element.Key + " 的命名区域指向不存在的工作表 " + sheet + "。"));
            return;
        }

        var hasImage = ExcelDrawingHelper.GetBlipEmbed(ExcelDrawingHelper.FindAnchor(worksheetPart, start.Col - 1, start.Row - 1)) is not null;
        if (!hasImage)
        {
            issues.Add(WrongType(element.Key, "图片元素 " + element.Key + " 的锚定格内没有图片（应锚定占位图）。"));
        }
    }

    private static void CheckTableElement(
        TableElement element,
        WorkbookPart workbookPart,
        IReadOnlyList<NamedRangeMatch> names,
        List<TemplateValidationIssue> issues)
    {
        var missingColumns = new List<string>();
        var rows = new List<int>();
        foreach (var column in element.Columns)
        {
            var name = ExcelNamedRangeLocator.TableColumnName(element.Key, column.Key);
            var match = names.FirstOrDefault(n => n.Name == name);
            if (match is null)
            {
                missingColumns.Add(column.Key);
            }
            else
            {
                var (_, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
                rows.Add(start.Row);
            }
        }

        var requiredKeys = element.Columns.Where(c => c.Required).Select(c => c.Key).ToHashSet(StringComparer.Ordinal);
        var missingRequired = missingColumns.Where(requiredKeys.Contains).ToList();
        var hasCompleteRow = rows.Count > 0 && rows.Distinct().Count() == 1;

        if (missingRequired.Count > 0 || !hasCompleteRow)
        {
            issues.Add(Missing(
                element.Key,
                "缺少完整表格行模板 " + element.Key + "（" + element.DisplayName + "）："
                + (missingRequired.Count > 0
                    ? "缺少必填列 " + string.Join(", ", missingRequired) + "。"
                    : "必填列未出现在同一示例行内。"),
                element.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning));
        }
    }

    /// <summary>按工作表名解析 WorksheetPart（定义名引用里的 Sheet 名）。</summary>
    internal static WorksheetPart? ResolveWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        if (workbookPart.Workbook?.Sheets is not { } sheets)
        {
            return null;
        }

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            if (sheet.Name?.Value == sheetName && sheet.Id?.Value is { } id)
            {
                return workbookPart.GetPartById(id) as WorksheetPart;
            }
        }

        return null;
    }

    private static TemplateValidationIssue Missing(
        string key,
        string message,
        TemplateValidationSeverity severity = TemplateValidationSeverity.Error)
        => new()
        {
            Code = TemplateValidationIssueCode.Missing,
            Key = key,
            Message = message,
            Severity = severity,
        };

    private static TemplateValidationIssue WrongType(string key, string message)
        => new()
        {
            Code = TemplateValidationIssueCode.WrongType,
            Key = key,
            Message = message,
        };

    private static TemplateValidationResult Invalid(string message)
        => new()
        {
            Issues =
            [
                new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Invalid,
                    Message = message,
                },
            ],
        };
}
