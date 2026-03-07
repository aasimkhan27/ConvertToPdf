using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpreadsheetToPdf
{
    internal sealed class ExcelPdfConverter : IDisposable
    {
        private static readonly object Missing = Type.Missing;

        private Excel.Application _excelApplication;
        private bool _disposed;

        public ExcelPdfConverter()
        {
            try
            {
                _excelApplication = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false,
                    EnableEvents = false,
                    AskToUpdateLinks = false,
                    UserControl = false
                };
            }
            catch (COMException ex)
            {
                throw new ExcelNotInstalledException(
                    "Microsoft Excel Interop could not be started. Ensure Microsoft Excel is installed on this machine.",
                    ex);
            }
        }

        public void Convert(ConversionRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            SpreadsheetFileType fileType = FileTypeDetector.Detect(request.InputPath);
            Console.WriteLine($"Detected input format: {fileType}");

            Excel.Workbook workbook = null;
            try
            {
                workbook = fileType == SpreadsheetFileType.Csv
                    ? OpenCsvWorkbook(request.InputPath)
                    : OpenWorkbook(request.InputPath);

                if (fileType == SpreadsheetFileType.Csv)
                {
                    ApplyMinimalCsvLayoutFixes(workbook);
                }

                ExportWorkbook(workbook, request.OutputPdfPath, request.WorksheetName);
            }
            catch (WorksheetNotFoundException)
            {
                throw;
            }
            catch (COMException ex)
            {
                throw new ExcelInteropException("COM error while processing the spreadsheet in Excel.", ex);
            }
            finally
            {
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch
                    {
                        // Ignore shutdown errors while we are already cleaning up.
                    }

                    ReleaseComObject(workbook);
                }
            }
        }

        private Excel.Workbook OpenWorkbook(string inputPath)
        {
            return _excelApplication.Workbooks.Open(
                inputPath,
                UpdateLinks: 0,
                ReadOnly: true,
                Format: Missing,
                Password: Missing,
                WriteResPassword: Missing,
                IgnoreReadOnlyRecommended: true,
                Origin: Missing,
                Delimiter: Missing,
                Editable: false,
                Notify: false,
                Converter: Missing,
                AddToMru: false,
                Local: true,
                CorruptLoad: Missing);
        }

        private Excel.Workbook OpenCsvWorkbook(string inputPath)
        {
            CsvImportSettings csvSettings = CsvImportSettings.Detect(inputPath);
            Console.WriteLine($"Detected CSV delimiter: '{csvSettings.Delimiter}'");

            _excelApplication.Workbooks.OpenText(
                Filename: inputPath,
                Origin: Excel.XlPlatform.xlWindows,
                StartRow: 1,
                DataType: Excel.XlTextParsingType.xlDelimited,
                TextQualifier: Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
                ConsecutiveDelimiter: false,
                Tab: csvSettings.Delimiter == '\t',
                Semicolon: csvSettings.Delimiter == ';',
                Comma: csvSettings.Delimiter == ',',
                Space: false,
                Other: csvSettings.UseOtherDelimiter,
                OtherChar: csvSettings.UseOtherDelimiter
                    ? csvSettings.Delimiter.ToString(CultureInfo.InvariantCulture)
                    : Missing,
                FieldInfo: Missing,
                DecimalSeparator: csvSettings.DecimalSeparator,
                ThousandsSeparator: csvSettings.ThousandSeparator,
                TrailingMinusNumbers: true,
                Local: true);

            Excel.Workbook activeWorkbook = _excelApplication.ActiveWorkbook;
            if (activeWorkbook == null)
            {
                throw new ExcelInteropException("Excel could not open the CSV workbook.", null);
            }

            return activeWorkbook;
        }

        private static void ApplyMinimalCsvLayoutFixes(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            Excel.Worksheet sheet = null;
            Excel.Range usedRange = null;
            Excel.Range columns = null;

            try
            {
                sheet = workbook.Worksheets[1] as Excel.Worksheet;
                if (sheet == null)
                {
                    return;
                }

                usedRange = sheet.UsedRange;
                columns = usedRange?.EntireColumn;

                // CSV has no intrinsic column widths. Auto-fit columns once to avoid clipped values.
                columns?.AutoFit();
            }
            finally
            {
                ReleaseComObject(columns);
                ReleaseComObject(usedRange);
                ReleaseComObject(sheet);
            }
        }

        private void ExportWorkbook(Excel.Workbook workbook, string outputPath, string worksheetName)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                Console.WriteLine("Exporting all worksheets to PDF...");
                workbook.ExportAsFixedFormat(
                    Type: Excel.XlFixedFormatType.xlTypePDF,
                    Filename: outputPath,
                    Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: true,
                    IgnorePrintAreas: false,
                    OpenAfterPublish: false);
                return;
            }

            Console.WriteLine($"Exporting worksheet '{worksheetName}' to PDF...");
            Excel.Worksheet sheet = FindWorksheet(workbook, worksheetName);
            if (sheet == null)
            {
                throw new WorksheetNotFoundException($"Worksheet '{worksheetName}' was not found in the workbook.");
            }

            try
            {
                sheet.ExportAsFixedFormat(
                    Type: Excel.XlFixedFormatType.xlTypePDF,
                    Filename: outputPath,
                    Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: true,
                    IgnorePrintAreas: false,
                    OpenAfterPublish: false);
            }
            finally
            {
                ReleaseComObject(sheet);
            }
        }

        private static Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string worksheetName)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                int sheetCount = sheets.Count;

                for (int i = 1; i <= sheetCount; i++)
                {
                    Excel.Worksheet currentSheet = null;
                    bool isMatch = false;

                    try
                    {
                        currentSheet = sheets[i] as Excel.Worksheet;
                        isMatch = currentSheet != null &&
                            string.Equals(currentSheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase);

                        if (isMatch)
                        {
                            return currentSheet;
                        }
                    }
                    finally
                    {
                        if (currentSheet != null && !isMatch)
                        {
                            ReleaseComObject(currentSheet);
                        }
                    }
                }

                return null;
            }
            finally
            {
                ReleaseComObject(sheets);
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            try
            {
                if (_excelApplication != null)
                {
                    _excelApplication.DisplayAlerts = false;
                    _excelApplication.ScreenUpdating = false;
                    _excelApplication.EnableEvents = false;
                    _excelApplication.Quit();
                }
            }
            finally
            {
                ReleaseComObject(_excelApplication);
                _excelApplication = null;
                _disposed = true;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch
            {
                // Ignore cleanup exceptions.
            }
        }
    }
}
