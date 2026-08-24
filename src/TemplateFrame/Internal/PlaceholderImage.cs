namespace TemplateFrame.Internal;

/// <summary>内置占位图（Word / Excel 插件共用）：未提供自定义占位图时的浅灰棋盘 240x120 PNG。</summary>
internal static class PlaceholderImage
{
    private const string PngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAPAAAAB4CAIAAABD1OhwAAACAUlEQVR4nO3asQnAMBAEwe+/KdfhbpQ6FQZjLfMFDBJseHNv3rV5fP6X/vztQXz+G1/Q/JQvaH7KFzQ/5Quan/IFzU/5guanfEHzU76g+Slf0PyUL2h+yhc0P+ULmp/yBc1P+YLmp/w5/QN8/vMEzU/5guanfEHzU76g+Slf0PyUL2h+yhc0P+ULmp/yBc1P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lG/jzU76g+Slf0PyUL2h+yhc0P+ULmp/yBc1P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8A39+yhc0P+ULmp/yBc1P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JRv4M9P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IN/PkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IFzU/5guanfEHzU76BPz/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IFzU/5guanfEHzU76g+Slf0PyUL2h+yjfw56d8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IFzU/5C7iGFURrlyOlAAAAAElFTkSuQmCC";

    /// <summary>
    /// 加载占位图：<paramref name="placeholderPath"/> 非空则读取该文件（扩展名按魔数识别，文件不存在直接抛），
    /// 否则返回内置 PNG。
    /// </summary>
    public static (byte[] Bytes, string Extension) Load(string? placeholderPath)
    {
        if (!string.IsNullOrWhiteSpace(placeholderPath))
        {
            var bytes = File.ReadAllBytes(placeholderPath);
            return (bytes, ImageTypeDetector.DetectExtension(bytes));
        }

        return (Convert.FromBase64String(PngBase64), "png");
    }
}
