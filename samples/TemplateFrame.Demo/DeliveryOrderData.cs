namespace TemplateFrame.Demo;

/// <summary>送货单明细行（序号由业务侧填写；实收数量/批次号/仓库在收货前为空）。</summary>
public sealed record DeliveryOrderLine
{
    public int RowNo { get; init; }

    public string MaterialCode { get; init; } = string.Empty;

    public string MaterialName { get; init; } = string.Empty;

    public string Unit { get; init; } = string.Empty;

    public decimal PlanQty { get; init; }

    /// <summary>实收数量（收货前为空）。</summary>
    public decimal? ActualQty { get; init; }

    /// <summary>批次号（收货前为空）。</summary>
    public string? BatchNo { get; init; }

    /// <summary>供应商批次号（计划送货时已知）。</summary>
    public string? SupplierBatchNo { get; init; }

    /// <summary>仓库（收货前为空）。</summary>
    public string? Warehouse { get; init; }
}

/// <summary>送货单数据（收货前：实际到货日期/收货人/实收数量/批次号/仓库为空；收货后补齐）。</summary>
public sealed record DeliveryOrderData
{
    public string Supplier { get; init; } = string.Empty;

    public string No { get; init; } = string.Empty;

    /// <summary>二维码内容（填充时由 Demo 用 QRCoder 生成 PNG 填进页眉二维码控件）。</summary>
    public string QrContent { get; init; } = string.Empty;

    public DateTime OrderDate { get; init; }

    public string OrderBy { get; init; } = string.Empty;

    public string Remark { get; init; } = string.Empty;

    public DateTime PlanDeliveryDate { get; init; }

    /// <summary>实际到货日期（收货前为空）。</summary>
    public DateTime? ActualArrivalDate { get; init; }

    /// <summary>收货人（收货前为空）。</summary>
    public string Receiver { get; init; } = string.Empty;

    public IReadOnlyList<DeliveryOrderLine> Lines { get; init; } = [];
}