using System;
using System.IO;

namespace SpreadsheetToPdf
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.WriteLine("SpreadsheetToPdf - Excel Interop PDF converter");
            Console.WriteLine("------------------------------------------------");

            if (!CliArguments.TryParse(args, out ConversionRequest request, out string parseError))
            {
                Console.Error.WriteLine(parseError);
                Console.WriteLine();
                CliArguments.PrintUsage();
                return ExitCodes.InvalidArguments;
            }

            try
            {
                string outputDirectory = Path.GetDirectoryName(request.OutputPdfPath);
                if (!string.IsNullOrWhiteSpace(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                using (var converter = new ExcelPdfConverter())
                {
                    converter.Convert(request);
                }

                Console.WriteLine("Conversion completed successfully.");
                return ExitCodes.Success;
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"Input file not found: {ex.FileName ?? ex.Message}");
                return ExitCodes.FileNotFound;
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.Error.WriteLine($"Access denied: {ex.Message}");
                return ExitCodes.AccessDenied;
            }
            catch (WorksheetNotFoundException ex)
            {
                Console.Error.WriteLine(ex.Message);
                return ExitCodes.WorksheetNotFound;
            }
            catch (InvalidDataException ex)
            {
                Console.Error.WriteLine($"Invalid format: {ex.Message}");
                return ExitCodes.InvalidFormat;
            }
            catch (ExcelNotInstalledException ex)
            {
                Console.Error.WriteLine(ex.Message);
                return ExitCodes.ExcelNotInstalled;
            }
            catch (ExcelInteropException ex)
            {
                Console.Error.WriteLine($"Excel automation error: {ex.Message}");
                return ExitCodes.ExcelInteropError;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected failure: {ex.Message}");
                return ExitCodes.UnexpectedError;
            }
        }
    }
}
