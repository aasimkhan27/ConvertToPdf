# ConvertToPdf

Solution contains **two projects**:

1. **`SpreadsheetToPdf.Core`** (Class Library, .NET Framework 4.8.1)
   - Contains spreadsheet-to-PDF conversion logic.
   - Primary converter: Excel Interop (`Microsoft.Office.Interop.Excel`) for highest fidelity.
   - Optional fallback for `.xlsx`: ClosedXML + PdfSharp (lower fidelity).

2. **`SpreadsheetToPdf`** (**plain ASP.NET Web API** project, .NET Framework 4.8.1)
   - IIS/IIS Express-hosted Web API project (not self-host).
   - References `SpreadsheetToPdf.Core` and exposes HTTP endpoints.

## API endpoints
- `GET /api/conversion/health`
- `POST /api/conversion/pdf` (`multipart/form-data`)
  - `file` (required)
  - `worksheetName` (optional)

## Build and run (Visual Studio 2022)
1. Open `SpreadsheetToPdf.sln`.
2. Restore NuGet packages.
3. Set startup project to **`SpreadsheetToPdf`**.
4. Run with **IIS Express** (F5).
5. Call `<base-url>/api/conversion/health`.

## Notes
- Keep **Excel Interop** installed on the host machine for best PDF fidelity.
- Fallback mode is lower fidelity and is used only for `.xlsx` when Excel is unavailable.
