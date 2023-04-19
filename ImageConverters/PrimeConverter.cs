using ImageConverters.Svg;
using System;
using System.Drawing;
using System.IO;
using CropRectangle = System.Drawing.Rectangle;

namespace ImageConverters
{
    public static class PrimeConverter
    {
        private const int MAXIMUM_SVG_PNG_SIZE = 800;
        private const int CROP_FRAME_SIZE = 15;

        public static string SvgToPng(string svgData, int size = 300, int offset = 64) => SvgToPng(svgData, size, size, offset);

        public static string SvgToPng(string svgData, int width = 300, int height = 300, int offset = 64)
        {
            if (width > MAXIMUM_SVG_PNG_SIZE || height > MAXIMUM_SVG_PNG_SIZE)
                throw new InvalidOperationException();

            using (var stream = SvgConverter.ToBitmapStream(svgData, width, height, offset))
            {
                var path = $"{Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())}.png";

                using (var bitmap = new Bitmap(stream))
                {
                    var rectangle = new CropRectangle(CROP_FRAME_SIZE, CROP_FRAME_SIZE, bitmap.Width - CROP_FRAME_SIZE, bitmap.Height - CROP_FRAME_SIZE);

                    using (var croppedBitmap = bitmap.Clone(rectangle, bitmap.PixelFormat))
                        croppedBitmap.Save(path);
                }

                return path;
            }
        }
    }
}