namespace TemplateFrame.Data;

/// <summary>The plugin-agnostic, weakly-typed data shape (scalar values plus table rows).</summary>
/// <remarks>
/// 数据形状：与插件无关的弱类型数据容器，替代 TemplateFiller 的 ISource 路径反射。
/// Word 内容控件 tag 是扁平键，嵌套对象在业务服务映射时展平（如 Customer.Name → tag CustomerName）。
/// </remarks>
public sealed class FillData
{
    /// <summary>Scalar fields: key (content-control tag) → value.</summary>
    public IReadOnlyDictionary<string, object?> Values { get; init; }
        = new Dictionary<string, object?>();

    /// <summary>Table data: table key → row list (each row maps column key → value).</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> Tables { get; init; }
        = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();
}
