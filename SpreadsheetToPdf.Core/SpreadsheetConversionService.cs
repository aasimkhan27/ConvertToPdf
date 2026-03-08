using System;
using System.IO;

namespace SpreadsheetToPdf.Core
{
    public sealed class SpreadsheetConversionService
    {
        public ConversionResult ConvertToPdf(string inputFilePath, string worksheetName)
        {
            if (string.IsNullOrWhiteSpace(inputFilePath))
            {
                throw new ArgumentException("Input file path is required.", nameof(inputFilePath));
            }

            string tempPdfPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
            var request = new ConversionRequest(inputFilePath, tempPdfPath, worksheetName);

            bool usedFallback = false;
            string message = "Converted using Excel Interop.";

            try
            {
                using (var converter = new ExcelPdfConverter())
                {
                    converter.Convert(request);
                }
            }
            catch (ExcelNotInstalledException)
            {
                SpreadsheetFileType inputType = FileTypeDetector.Detect(inputFilePath);
                if (inputType != SpreadsheetFileType.Xlsx)
                {
                    throw;
                }

                usedFallback = true;
                message = "Converted using fallback mode (lower fidelity than Excel Interop).";
                new XlsxFallbackPdfConverter().Convert(request);
            }

            byte[] pdfBytes = File.ReadAllBytes(tempPdfPath);
            string outputFileName = Path.GetFileNameWithoutExtension(inputFilePath) + ".pdf";

            TryDeleteFile(tempPdfPath);

            return new ConversionResult
            {
                FileName = outputFileName,
                ContentType = "application/pdf",
                Content = pdfBytes,
                UsedFallback = usedFallback,
                Message = message
            };
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }
    }
}
