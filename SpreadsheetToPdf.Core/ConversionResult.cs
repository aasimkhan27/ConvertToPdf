namespace SpreadsheetToPdf.Core
{
    public sealed class ConversionResult
    {
        public string FileName { get; set; }

        public string ContentType { get; set; }

        public byte[] Content { get; set; }

        public bool UsedFallback { get; set; }

        public string Message { get; set; }
    }
}
