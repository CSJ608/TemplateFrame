using QRCoder;

namespace TemplateFrame.Demo.Excel.AutoMapping;

/// <summary>Demo 侧二维码生成（QRCoder，属于业务侧能力，不进框架）。</summary>
internal static class QrCodeGenerator
{
    /// <summary>生成二维码 PNG 字节（内容、容错级别 M）。</summary>
    public static byte[] CreatePng(string content, int pixelsPerModule = 8)
    {
        using var generator = new QRCodeGenerator();
        var qrData = generator.CreateQrCode(content, QRCodeGenerator.ECCLevel.M);
        using var qr = new PngByteQRCode(qrData);
        return qr.GetGraphic(pixelsPerModule);
    }
}
