namespace TemplateFrame.Contract;

/// <summary>A table element (detail area) — a sample row with content controls, cloned per row on fill.</summary>
/// <remarks>
/// 表格元素：明细区。模板里放一行"示例行"（每格一个内容控件），填充时按行克隆。
/// 行的单元格 tag 使用列 <see cref="Columns"/> 的 Key，因此列 Key 也必须在文档内全局唯一。
/// </remarks>
public sealed record TableElement : TemplateElement
{
    /// <summary>Row-template fields (one TextElement per cell).</summary>
    public IReadOnlyList<TextElement> Columns { get; init; } = [];
}
