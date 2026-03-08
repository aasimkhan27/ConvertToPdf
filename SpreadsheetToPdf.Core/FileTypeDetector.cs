using System;
using System.IO;

namespace SpreadsheetToPdf.Core
{
    internal enum SpreadsheetFileType
    {
        Xlsx,
        Xls,
        Csv
    }

    internal static class FileTypeDetector
    {
        public static SpreadsheetFileType Detect(string inputPath)
        {
            string extension = Path.GetExtension(inputPath);

            if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return SpreadsheetFileType.Xlsx;
            }

            if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase))
            {
                return SpreadsheetFileType.Xls;
            }

            if (string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase))
            {
                return SpreadsheetFileType.Csv;
            }

            throw new InvalidDataException($"Unsupported input extension '{extension}'. Supported: .xlsx, .xls, .csv.");
        }
    }
}
