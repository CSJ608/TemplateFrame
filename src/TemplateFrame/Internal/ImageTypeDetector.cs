namespace TemplateFrame.Internal;

internal static class ImageTypeDetector
{
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

    public static string ToImagePartType(string extension)
        => extension switch
        {
            "jpg" or "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" => "image/tiff",
            _ => "image/png",
        };

    public static string DetectContentType(byte[] bytes)
        => ToImagePartType(DetectExtension(bytes));
}
