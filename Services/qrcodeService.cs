using PdfSharp.Drawing;
using ZXing;
using ZXing.Common;

namespace QueueAppManager.Service
{
    public static class qrcodeService
    {
        public static BitMatrix Encode(string qrCodeContent, int size = 50)
        {
            var qrCodeWriter = new BarcodeWriterGeneric()
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new EncodingOptions { Width = size, Height = size, },
            };

            BitMatrix bitMatrix = qrCodeWriter.Encode(qrCodeContent);
            return bitMatrix;
        }

        public static void DrawQrCode(this XGraphics graphics, XPoint upperLeftCorner, BitMatrix bitMatrix, double pixelSize = 1.0)
        {
            for (int x = 0; x < bitMatrix.Width; x++)
            {
                for (int y = 0; y < bitMatrix.Height; y++)
                {
                    var pixelLocation = new XPoint(upperLeftCorner.X + (x * pixelSize), upperLeftCorner.Y + (y * pixelSize));
                    var rectangle = new XRect(pixelLocation, new XSize(pixelSize, pixelSize));
                    var color = bitMatrix[x, y] ? XBrushes.Black : XBrushes.White;

                    graphics.DrawRectangle(color, rectangle);
                }
            }
        }
    }
}
