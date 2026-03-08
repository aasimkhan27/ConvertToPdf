namespace SpreadsheetToPdf
{
    internal sealed class ConversionRequest
    {
        public ConversionRequest(string inputPath, string outputPdfPath, string worksheetName)
        {
            InputPath = inputPath;
            OutputPdfPath = outputPdfPath;
            WorksheetName = worksheetName;
        }

        public string InputPath { get; }

        public string OutputPdfPath { get; }

        public string WorksheetName { get; }

        public bool HasWorksheetName => !string.IsNullOrWhiteSpace(WorksheetName);
    }
}
