namespace TemplateFrame.Builder;

/// <summary>
/// 版式构建器的唯一通用契约：把组装结果写入输出流。
/// <para>English: The only shared contract of a layout builder — writes the assembled result to a stream.</para>
/// 具体插件的构建器（如 <c>WordTemplateBuilder</c>）在各自类型上直接暴露全部排版能力，
/// 业务服务声明 `TemplateService&lt;TData, TBuilder&gt;` 后，在 <c>BuildInitialTemplate()</c> 里直接用具体实例。
/// </summary>
public interface ITemplateBuilder
{
    /// <summary>把组装结果写入输出流。</summary>
    void Save(Stream target);
}