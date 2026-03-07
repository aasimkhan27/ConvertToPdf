using System;
using System.IO;

namespace SpreadsheetToPdf
{
    internal static class CliArguments
    {
        public static bool TryParse(string[] args, out ConversionRequest request, out string error)
        {
            request = null;
            error = string.Empty;

            if (args == null || args.Length < 2 || args.Length > 3)
            {
                error = "Invalid arguments. Expected 2 or 3 arguments.";
                return false;
            }

            string inputPath = Path.GetFullPath(args[0]);
            string outputPath = Path.GetFullPath(args[1]);
            string worksheetName = args.Length == 3 ? args[2]?.Trim() : null;

            if (string.IsNullOrWhiteSpace(inputPath))
            {
                error = "Input path cannot be empty.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                error = "Output PDF path cannot be empty.";
                return false;
            }

            if (!File.Exists(inputPath))
            {
                error = $"Input file does not exist: {inputPath}";
                return false;
            }

            if (!string.Equals(Path.GetExtension(outputPath), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                error = "Output file must use the .pdf extension.";
                return false;
            }

            request = new ConversionRequest(inputPath, outputPath, worksheetName);
            return true;
        }

        public static void PrintUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  SpreadsheetToPdf.exe \"<input-file>\" \"<output-pdf>\" [worksheet-name]");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  SpreadsheetToPdf.exe \"C:\\Files\\report.xlsx\" \"C:\\Files\\report.pdf\"");
            Console.WriteLine("  SpreadsheetToPdf.exe \"C:\\Files\\report.xls\" \"C:\\Files\\report.pdf\" \"Sheet1\"");
            Console.WriteLine("  SpreadsheetToPdf.exe \"C:\\Files\\data.csv\" \"C:\\Files\\data.pdf\"");
        }
    }
}
