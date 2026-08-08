using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;

namespace TemplateFrame.Demo.Excel.Simple.I18n;

/// <summary>
/// 物料基础数据场景服务：依赖契约（单个表格：编码 / 名称 / 基本单位 / 包装规格 / 型号）。
/// 表格与列声明 <see cref="TemplateElement.DataPath"/> 后由基础包自动映射，
/// 无需手写 MapToData / MapFromData，即可获得强类型 BuildTemplate / Validate / Fill / Parse。
/// </summary>
public sealed class MaterialsTemplateService : SimpleExcelTemplateService<MaterialsData>
{
    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "Materials",
            Version = "1.0",
            Elements =
            [
                new TableElement
                {
                    Key = "Materials",
                    DisplayName = "物料清单",
                    DataPath = "Items",
                    Columns =
                    [
                        new TextElement { Key = "编码", DisplayName = "编码", DataPath = "Code", Required = true },
                        new TextElement { Key = "名称", DisplayName = "名称", DataPath = "Name", Required = true },
                        new TextElement { Key = "基本单位", DisplayName = "基本单位", DataPath = "Unit" },
                        new TextElement { Key = "包装规格", DisplayName = "包装规格", DataPath = "Package" },
                        new TextElement { Key = "型号", DisplayName = "型号", DataPath = "Model" },
                    ],
                },
            ],
        };
}

/// <summary>
/// 消息层 i18n 演示服务（迭代 12）：契约**故意缺「型号」列**（只有 4 列），
/// 用它生成的模板拿去对完整契约（5 列）做 <c>Validate</c>，报 Missing 的消息
/// 随 <c>CurrentUICulture</c> 中英切换（MessageKey / MessageArgs 稳定）。
/// </summary>
public sealed class MaterialsMissingModelTemplateService : SimpleExcelTemplateService<MaterialsData>
{
    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "MaterialsMissingModel",
            Version = "1.0",
            Elements =
            [
                new TableElement
                {
                    Key = "Materials",
                    DisplayName = "物料清单",
                    DataPath = "Items",
                    Columns =
                    [
                        new TextElement { Key = "编码", DisplayName = "编码", DataPath = "Code", Required = true },
                        new TextElement { Key = "名称", DisplayName = "名称", DataPath = "Name", Required = true },
                        new TextElement { Key = "基本单位", DisplayName = "基本单位", DataPath = "Unit" },
                        new TextElement { Key = "包装规格", DisplayName = "包装规格", DataPath = "Package" },
                    ],
                },
            ],
        };
}