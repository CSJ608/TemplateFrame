namespace TemplateFrame.Data;

/// <summary>
/// 数据形状：与插件无关的弱类型数据容器，替代 TemplateFiller 的 ISource 路径反射。
/// Word 内容控件 tag 是扁平键，嵌套对象在业务服务映射时展平（如 Customer.Name → tag CustomerName）。
/// </summary>
public sealed class FillData
{
    /// <summary>标量字段：Key（内容控件 tag）→ 值。</summary>
    public IReadOnlyDictionary<string, object?> Values { get; init; }
        = new Dictionary<string, object?>();

    /// <summary>表格数据：表格 Key → 行集合（每行是列 Key → 值）。</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> Tables { get; init; }
        = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();
}
