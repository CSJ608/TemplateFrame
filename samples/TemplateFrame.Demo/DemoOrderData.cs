namespace TemplateFrame.Demo;

/// <summary>Demo 单据行（明细行）。</summary>
public sealed record DemoOrderLine
{
    public string MaterialCode { get; init; } = string.Empty;
    public string MaterialName { get; init; } = string.Empty;
    public decimal Quantity { get; init; }
}

/// <summary>Demo 单据数据（示例场景服务使用的强类型数据）。</summary>
public sealed record DemoOrderData
{
    public string OrderNo { get; init; } = string.Empty;
    public string CustomerName { get; init; } = string.Empty;
    public DateTime OrderDate { get; init; }
    public decimal TotalAmount { get; init; }
    public IReadOnlyList<DemoOrderLine> Lines { get; init; } = [];
}
