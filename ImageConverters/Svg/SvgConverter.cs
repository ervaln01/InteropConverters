using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ImageConverters.Svg.Interfaces;
using ImageConverters.Svg.Structs;
using mshtml;
using Draw = System.Drawing;

namespace ImageConverters.Svg
{
    public class SvgConverter
    {
        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;
        private const int HIMETRIC_INCH = 2540;
        private const int SPI_GETWORKAREA = 48;

        private const string HTML_FORMAT = @"<html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=11"" /></head><body>{0}</body></html>";

        public static Stream ToBitmapStream(string content, int width, int height, int offset = 130)
        {
            var rectangle = GetRectangle();

            HTMLDocument htmlDocument = null;
            IHTMLDocument2 htmlDocument2 = null;
            IOleObject oleObject = null;
            IViewObject viewObject = null;

            try
            {
                htmlDocument = new HTMLDocument();
                htmlDocument2 = (IHTMLDocument2)htmlDocument;

                oleObject = (IOleObject)htmlDocument2;
                SetDocumentSize(rectangle, oleObject);

                htmlDocument2.write(string.Format(HTML_FORMAT, content));
                htmlDocument2.close();

                var stream = new MemoryStream();

                using (var bitmap = new Draw.Bitmap(width + offset, height + offset))
                {
                    viewObject = (IViewObject)htmlDocument2;
                    FillBitmap(bitmap, viewObject, rectangle);
                    bitmap.Save(stream, ImageFormat.Bmp);
                }

                return stream;
            }
            finally
            {
                ReleaseComObjects(htmlDocument2, oleObject, viewObject, htmlDocument);
            }
        }

        private static void FillBitmap(Draw.Bitmap bitmap, IViewObject viewObject, Rectangle rcClient)
        {
            using (var graphics = Draw.Graphics.FromImage(bitmap))
            {
                var deviceContextHandler = graphics.GetHdc();
                var hresult = viewObject.Draw(
                    dwDrawAspect: (int)DVASPECT.DVASPECT_CONTENT,
                    hdcDraw: deviceContextHandler,
                    lprcBounds: ref rcClient,
                    lindex: -1,
                    dwContinue: 0,
                    pvAspect: IntPtr.Zero,
                    ptd: IntPtr.Zero,
                    hdcTargetDev: IntPtr.Zero,
                    lprcWBounds: IntPtr.Zero,
                    pfnContinue: IntPtr.Zero);

                ThrownIfError(hresult != 0, hresult);

                graphics.ReleaseHdc(deviceContextHandler);
            }
        }

        private static Rectangle GetRectangle()
        {
            var rectangle = new Rectangle();

            var isCorrect = SystemParametersInfo(SPI_GETWORKAREA, 0, ref rectangle, 0);

            if (!isCorrect)
            {
                rectangle.Bottom = 480;
                rectangle.Right = 640;
            }

            return rectangle;
        }

        private static void SetDocumentSize(Rectangle rcClient, IOleObject oleObject)
        {
            var deviceContext = GetDC(IntPtr.Zero);

            var width = (int)(rcClient.Right - rcClient.Left);
            var height = (int)(rcClient.Bottom - rcClient.Top);

            var size = new Size
            {
                X = (uint)MulDiv(width, HIMETRIC_INCH, GetDeviceCaps(deviceContext, LOGPIXELSX)),
                Y = (uint)MulDiv(height, HIMETRIC_INCH, GetDeviceCaps(deviceContext, LOGPIXELSY))
            };

            var hResult = oleObject.SetExtent((int)DVASPECT.DVASPECT_CONTENT, ref size);

            ThrownIfError(hResult != 0, hResult);
        }

        private static void ThrownIfError(bool condition, int errorCode)
        {
            if (condition)
                throw Marshal.GetExceptionForHR(errorCode);
        }

        private static void ReleaseComObjects(params object[] objects) =>
            objects.Where(o => o != null).ToList().ForEach(o => Marshal.ReleaseComObject(o));

        private static int MulDiv(int number, int numerator, int denominator) =>
            (int)((long)number * numerator / denominator);

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool SystemParametersInfo(int nAction, int nParam, ref Rectangle rc, int nUpdate);
    }
}