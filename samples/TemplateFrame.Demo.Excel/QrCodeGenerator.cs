using Net.Codecrete.QrCodeGenerator;

namespace TemplateFrame.Demo.Excel;

/// <summary>Demo 侧二维码生成（纯托管、跨平台 PNG，属于业务侧能力，不进框架）。</summary>
internal static class QrCodeGenerator
{
    /// <summary>生成二维码 PNG 字节（内容、容错级别 M）。</summary>
    public static byte[] CreatePng(string content, int pixelsPerModule = 8)
    {
        var qr = QrCode.EncodeTextAdvanced(content, QrCode.Ecc.Medium, boostEcl: false);
        return qr.ToPngBitmap(4, pixelsPerModule);
    }
}
