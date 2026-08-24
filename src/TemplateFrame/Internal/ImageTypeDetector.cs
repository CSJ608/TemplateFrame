namespace TemplateFrame.Internal;

/// <summary>内部图片类型探测（Word / Excel 插件共用）：按文件头魔数识别扩展名并映射 MIME。</summary>
internal static class ImageTypeDetector
{
    /// <summary>按文件头魔数识别图片扩展名（png / jpg / gif / bmp / tiff；无法识别回退 png）。</summary>
    public static string DetectExtension(byte[] bytes)
    {
        if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
        {
            return "png";
        }

        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
        {
            return "jpg";
        }

        if (bytes.Length >= 6 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
        {
            return "gif";
        }

        if (bytes.Length >= 2 && bytes[0] == 0x42 && bytes[1] == 0x4D)
        {
            return "bmp";
        }

        if (bytes.Length >= 4
            && ((bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2A && bytes[3] == 0x00)
                || (bytes[0] == 0x4D && bytes[1] == 0x4D && bytes[2] == 0x00 && bytes[3] == 0x2A)))
        {
            return "tiff";
        }

        return "png";
    }

    /// <summary>扩展名 → OpenXML 图片 MIME（image/png 兜底）。</summary>
    public static string ToImagePartType(string extension)
        => extension switch
        {
            "jpg" or "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" => "image/tiff",
            _ => "image/png",
        };

    /// <summary>按字节识别图片 MIME（Excel drawing 等使用）。</summary>
    public static string DetectContentType(byte[] bytes)
        => ToImagePartType(DetectExtension(bytes));
}
