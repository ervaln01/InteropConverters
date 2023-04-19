using System;
using System.IO;
using DocumentConverters.Pdf;

namespace DocumentConverters
{
    public static class PrimeConverter
    {
        public static string DocxToPdf(string docxFileName)
        {
            var pdfFileName = docxFileName.Replace(".docx", ".pdf");

            if (!File.Exists(docxFileName))
                throw new InvalidOperationException("Input file not exists");

            var result = PdfConverter.ConvertFromDocx(docxFileName, pdfFileName);

            if (!result.CorrectConvert)
                throw new InvalidOperationException(result.ErrorMessage);

            return pdfFileName;
        }
    }
}
