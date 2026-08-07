namespace TemplateFrame.Demo.Word.AutoMapping;

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

/// <summary>
/// 送货单数据（自动映射版）：与手写映射版内容一致，区别在于——
/// 图片字节由数据直接携带（<see cref="LogoBytes"/> / <see cref="QrBytes"/>），
/// 不再在 MapToData 里读文件 / 生成二维码；契约元素声明 DataPath 后由框架自动映射。
/// </summary>
public sealed record DeliveryOrderData
{
    public string Supplier { get; init; } = string.Empty;

    public string No { get; init; } = string.Empty;

    public DateTime OrderDate { get; init; }

    public string OrderBy { get; init; } = string.Empty;

    public string Remark { get; init; } = string.Empty;

    public DateTime PlanDeliveryDate { get; init; }

    /// <summary>实际到货日期（收货前为空）。</summary>
    public DateTime? ActualArrivalDate { get; init; }

    /// <summary>收货人（收货前为空）。</summary>
    public string Receiver { get; init; } = string.Empty;

    /// <summary>公司 LOGO 图片字节（DataPath = Logo，填充页眉 LOGO 控件）。</summary>
    public byte[] LogoBytes { get; init; } = [];

    /// <summary>二维码图片字节（DataPath = QRCode，填充页眉二维码控件）。</summary>
    public byte[] QrBytes { get; init; } = [];

    public IReadOnlyList<DeliveryOrderLine> Lines { get; init; } = [];
}