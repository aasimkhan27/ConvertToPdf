namespace SpreadsheetToPdf
{
    internal static class ExitCodes
    {
        public const int Success = 0;
        public const int InvalidArguments = 1;
        public const int FileNotFound = 2;
        public const int InvalidFormat = 3;
        public const int ExcelNotInstalled = 4;
        public const int WorksheetNotFound = 5;
        public const int AccessDenied = 6;
        public const int ExcelInteropError = 7;
        public const int UnexpectedError = 99;
    }
}
