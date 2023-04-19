using System;
using Microsoft.Office.Interop.Word;

namespace DocumentConverters.Pdf
{
    public static class PdfConverter
    {
        public static ConverterReport ConvertFromDocx(object sourcePath, object targetPath)
        {
            object targetType = WdExportFormat.wdExportFormatPDF;
            object paramMissing = Type.Missing;

            ApplicationClass wordApplication = null;
            Document wordDocument = null;

            try
            {
                wordApplication = new ApplicationClass();
                wordDocument = wordApplication.Documents.Open(ref sourcePath);
                wordDocument?.SaveAs(targetPath, targetType);
            }
            catch (Exception ex)
            {
                return new ConverterReport(false, ex.Message);
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }

                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
            }

            return new ConverterReport(true, string.Empty);
        }

    }
}
