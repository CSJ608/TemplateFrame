namespace TemplateFrame.Demo;

/// <summary>送货单明细行（行号由业务侧填写）。</summary>
public sealed record DeliveryOrderLine
{
    public int RowNo { get; init; }

    public string MaterialName { get; init; } = string.Empty;

    public decimal Qty { get; init; }

    public string Unit { get; init; } = string.Empty;
}

/// <summary>送货单数据（送货单 Demo 使用的强类型数据）。</summary>
public sealed record DeliveryOrderData
{
    public string Supplier { get; init; } = string.Empty;

    public string No { get; init; } = string.Empty;

    /// <summary>二维码内容（填充时由 Demo 用 QRCoder 生成 PNG 填进页眉二维码控件）。</summary>
    public string QrContent { get; init; } = string.Empty;

    public DateTime PrintTime { get; init; }

    public string Printer { get; init; } = string.Empty;

    public DateTime? ArrivalTime { get; init; }

    public string Receiver { get; init; } = string.Empty;

    public IReadOnlyList<DeliveryOrderLine> Lines { get; init; } = [];
}