using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SpreadsheetToPdf.Core
{
    internal sealed class XlsxFallbackPdfConverter
    {
        private const double Margin = 30;
        private const double TitleHeight = 20;
        private const double RowHeight = 18;
        private const double MinColumnWidth = 50;

        public void Convert(ConversionRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            using (var workbook = new XLWorkbook(request.InputPath))
            using (var document = new PdfDocument())
            {
                IReadOnlyList<IXLWorksheet> worksheets = SelectWorksheets(workbook, request.WorksheetName);

                foreach (IXLWorksheet worksheet in worksheets)
                {
                    RenderWorksheet(document, worksheet);
                }

                document.Save(request.OutputPdfPath);
            }
        }

        private static IReadOnlyList<IXLWorksheet> SelectWorksheets(XLWorkbook workbook, string worksheetName)
        {
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                return workbook.Worksheets.ToList();
            }

            IXLWorksheet worksheet = workbook.Worksheets.FirstOrDefault(ws =>
                string.Equals(ws.Name, worksheetName, StringComparison.OrdinalIgnoreCase));

            if (worksheet == null)
            {
                throw new WorksheetNotFoundException($"Worksheet '{worksheetName}' was not found in the workbook.");
            }

            return new[] { worksheet };
        }

        private static void RenderWorksheet(PdfDocument document, IXLWorksheet worksheet)
        {
            IXLRange usedRange = worksheet.RangeUsed();

            if (usedRange == null)
            {
                PdfPage emptyPage = document.AddPage();
                using (XGraphics graphics = XGraphics.FromPdfPage(emptyPage))
                {
                    DrawWorksheetTitle(graphics, worksheet.Name);
                    graphics.DrawString("(Empty worksheet)", new XFont("Arial", 9), XBrushes.DarkGray,
                        new XRect(Margin, Margin + TitleHeight + 5, emptyPage.Width - 2 * Margin, 20),
                        XStringFormats.TopLeft);
                }

                return;
            }

            int firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
            int lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
            int firstColumn = usedRange.RangeAddress.FirstAddress.ColumnNumber;
            int lastColumn = usedRange.RangeAddress.LastAddress.ColumnNumber;

            int columnCount = (lastColumn - firstColumn) + 1;

            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            DrawWorksheetTitle(gfx, worksheet.Name);

            double availableWidth = page.Width - (2 * Margin);
            double columnWidth = Math.Max(MinColumnWidth, availableWidth / Math.Max(1, columnCount));

            int rowIndex = firstRow;

            while (rowIndex <= lastRow)
            {
                double y = Margin + TitleHeight;
                double availableHeight = page.Height - Margin - y;
                int maxRowsOnPage = Math.Max(1, (int)Math.Floor(availableHeight / RowHeight));

                int endingRowOnPage = Math.Min(lastRow, rowIndex + maxRowsOnPage - 1);

                for (int currentRow = rowIndex; currentRow <= endingRowOnPage; currentRow++)
                {
                    double x = Margin;
                    bool isHeaderRow = currentRow == firstRow;

                    for (int currentColumn = firstColumn; currentColumn <= lastColumn; currentColumn++)
                    {
                        IXLCell cell = worksheet.Cell(currentRow, currentColumn);
                        string text = FormatCellText(cell);

                        DrawCell(gfx, x, y, columnWidth, RowHeight, text, isHeaderRow);
                        x += columnWidth;
                    }

                    y += RowHeight;
                }

                rowIndex = endingRowOnPage + 1;
                gfx.Dispose();

                if (rowIndex <= lastRow)
                {
                    page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    DrawWorksheetTitle(gfx, worksheet.Name + " (cont.)");
                }
            }
        }

        private static void DrawWorksheetTitle(XGraphics graphics, string worksheetName)
        {
            var titleFont = new XFont("Arial", 12, XFontStyle.Bold);
            graphics.DrawString($"Worksheet: {worksheetName}", titleFont, XBrushes.Black,
                new XRect(Margin, Margin - 10, 500, TitleHeight),
                XStringFormats.TopLeft);
        }

        private static void DrawCell(XGraphics graphics, double x, double y, double width, double height, string text, bool isHeader)
        {
            XBrush backgroundBrush = isHeader ? XBrushes.LightGray : XBrushes.White;
            graphics.DrawRectangle(backgroundBrush, x, y, width, height);
            graphics.DrawRectangle(XPens.Gray, x, y, width, height);

            XFont font = isHeader ? new XFont("Arial", 8, XFontStyle.Bold) : new XFont("Arial", 8, XFontStyle.Regular);
            string clipped = ClipToLength(text, 120);

            graphics.DrawString(clipped, font, XBrushes.Black,
                new XRect(x + 2, y + 2, width - 4, height - 4),
                XStringFormats.TopLeft);
        }

        private static string FormatCellText(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
            {
                return string.Empty;
            }

            // ClosedXML keeps formatted display values through GetFormattedString().
            return cell.GetFormattedString();
        }

        private static string ClipToLength(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= maxLength)
            {
                return value;
            }

            return value.Substring(0, maxLength - 3) + "...";
        }
    }
}
