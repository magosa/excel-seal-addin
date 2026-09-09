using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace SealAddIn.Services
{
    /// <summary>
    /// GDI+で印影(スタンプ)のラスター画像(PNG)を生成します。
    /// 現行Web版のSVG生成+Canvasラスタライズと同等の見た目を、ネイティブ描画で直接生成します。
    /// </summary>
    public static class StampImageRenderer
    {
        private const int RasterSize = 400; // 高解像度化のためのラスターサイズ(px)

        public static byte[] Render(string text, string shape, string colorHex)
        {
            var color = ColorTranslator.FromHtml(string.IsNullOrWhiteSpace(colorHex) ? "#cc0000" : colorHex);
            var chars = string.IsNullOrEmpty(text) ? new[] { "印" } : text.Select(c => c.ToString()).ToArray();

            using (var bitmap = new Bitmap(RasterSize, RasterSize, PixelFormat.Format32bppArgb))
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                g.Clear(Color.Transparent);

                var strokeWidth = RasterSize * 0.045f;
                using (var pen = new Pen(color, strokeWidth))
                {
                    var inset = strokeWidth;
                    var size = RasterSize - inset * 2;
                    if (shape == "rect")
                    {
                        g.DrawRectangle(pen, inset, inset, size, size);
                    }
                    else
                    {
                        g.DrawEllipse(pen, inset, inset, size, size);
                    }
                }

                var fontSize = chars.Length <= 2
                    ? RasterSize * 0.26f
                    : Math.Max(RasterSize * 0.16f, (RasterSize * 0.8f) / chars.Length);

                using (var font = new Font("MS Mincho", fontSize, GraphicsUnit.Pixel))
                using (var brush = new SolidBrush(color))
                using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                {
                    var lineHeight = fontSize * 1.05f;
                    var totalHeight = lineHeight * chars.Length;
                    var startY = RasterSize / 2f - totalHeight / 2f + lineHeight / 2f;

                    for (var i = 0; i < chars.Length; i++)
                    {
                        g.DrawString(chars[i], font, brush, new PointF(RasterSize / 2f, startY + i * lineHeight), format);
                    }
                }

                using (var ms = new MemoryStream())
                {
                    bitmap.Save(ms, ImageFormat.Png);
                    return ms.ToArray();
                }
            }
        }
    }
}
