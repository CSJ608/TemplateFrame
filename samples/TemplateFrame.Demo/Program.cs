using TemplateFrame.Validation;
using TemplateFrame.Word;

namespace TemplateFrame.Demo;

internal static class Program
{
    private static void Main(string[] args)
    {
        var service = new DemoOrderTemplateService();

        Console.WriteLine($"契约：{service.Contract.Name} v{service.Contract.Version}，元素 {service.Contract.Elements.Count} 个");
        foreach (var element in service.Contract.Elements)
        {
            Console.WriteLine($"  - {element.Key}（{element.DisplayName}）");
        }

        // 1) 生成初始模板（含内容控件 SDT）
        var outputPath = args.Length > 0
            ? args[0]
            : Path.Combine(Path.GetTempPath(), "TemplateFrame-Demo", "DemoOrder-template.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        using (var template = service.BuildInitialTemplateFile())
        {
            File.WriteAllBytes(outputPath, ((MemoryStream)template).ToArray());
        }

        Console.WriteLine($"\n已生成初始模板：{outputPath}");

        // 2) Validate 兜底：枚举控件 + 报告 Missing/WrongType/Ambiguous
        using (var templateStream = File.OpenRead(outputPath))
        {
            var result = service.Validate(templateStream);
        Console.WriteLine($"\nValidate：{(result.IsValid ? "通过" : "未通过")}");

            if (result is WordTemplateValidationResult wordResult)
            {
                Console.WriteLine($"SDT 清单（{wordResult.Sdts.Count} 个）：");
                foreach (var sdt in wordResult.Sdts)
                {
                    Console.WriteLine($"  - tag={sdt.Tag,-14} id={sdt.Id,-4} kind={sdt.Kind,-6} location={sdt.Location}");
                }
            }

            if (result.Issues.Count > 0)
            {
                Console.WriteLine("问题清单：");
                foreach (var issue in result.Issues)
                {
                    Console.WriteLine($"  - [{issue.Code}] {issue.Message}（{issue.Severity}）");
                }
            }
        }
    }
}
