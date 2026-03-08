namespace SpreadsheetToPdf.Models
{
    public sealed class ConversionResponseDto
    {
        public string FileName { get; set; }

        public string ContentType { get; set; }

        public byte[] Content { get; set; }

        public bool UsedFallback { get; set; }

        public string Message { get; set; }
    }
}
