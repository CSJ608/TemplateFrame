namespace TemplateFrame.Demo.Excel.Simple;

/// <summary>物料基础数据行（对应契约表格的列：编码 / 名称 / 基本单位 / 包装规格 / 型号）。</summary>
public sealed record MaterialLine
{
    public string Code { get; init; } = string.Empty;

    public string Name { get; init; } = string.Empty;

    public string Unit { get; init; } = string.Empty;

    public string Package { get; init; } = string.Empty;

    public string Model { get; init; } = string.Empty;
}

/// <summary>物料基础数据：表格行集合（对应契约里单个表格的 DataPath = Items）。</summary>
public sealed record MaterialsData
{
    public IReadOnlyList<MaterialLine> Items { get; init; } = [];
}