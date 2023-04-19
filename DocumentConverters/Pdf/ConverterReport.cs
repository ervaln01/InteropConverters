namespace DocumentConverters.Pdf
{
    public class ConverterReport
    {
        public bool CorrectConvert { get; }
        public string ErrorMessage { get; }
        public ConverterReport(bool correctConvert, string errorMessage)
        {
            CorrectConvert = correctConvert;
            ErrorMessage = errorMessage;
        }
    }
}