namespace TemplateFrame.Builder;

/// <summary>The only shared contract of a layout builder — writes the assembled result to a stream.</summary>
/// <remarks>
/// 具体插件的构建器（如 <c>WordTemplateBuilder</c>）在各自类型上直接暴露全部排版能力，
/// 业务服务声明 `TemplateService&lt;TData, TBuilder&gt;` 后，在 <c>BuildInitialTemplate()</c> 里直接用具体实例。
/// </remarks>
public interface ITemplateBuilder
{
    /// <summary>Writes the assembled result to the output stream.</summary>
    void Save(Stream target);
}
