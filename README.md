# ConvertToPdf - ASP.NET Web API (.NET Framework 4.8.1)

This project is now an **ASP.NET Web API** service (Windows, .NET Framework 4.8.1) for converting spreadsheet files to PDF.

## Conversion behavior
- **Primary path (preferred):** Microsoft Excel Interop (`ExportAsFixedFormat`) for highest fidelity.
- **Fallback path:** if Excel is not installed and input is `.xlsx`, the API uses:
  - ClosedXML to read worksheet data
  - PdfSharp to render a basic table PDF
  - lower fidelity than Excel Interop (data-focused output)

Supported inputs:
- `.xlsx`
- `.xls`
- `.csv`

Output:
- `.pdf`

## API endpoints
### Health check
- `GET /api/conversion/health`

### Convert to PDF
- `POST /api/conversion/pdf`
- Content type: `multipart/form-data`
- Form fields:
  - `file` (required): spreadsheet file
  - `worksheetName` (optional): specific worksheet name

Response:
- `200 OK` with PDF bytes (`application/pdf`)
- Response headers:
  - `X-Used-Fallback: true|false`
  - `X-Conversion-Message: ...`

## Example request (curl)
```bash
curl -X POST "http://localhost:port/api/conversion/pdf" \
  -F "file=@C:/Files/report.xlsx" \
  -F "worksheetName=Sheet1" \
  --output report.pdf
```

## Prerequisites
1. Windows
2. Visual Studio 2022 (or compatible)
3. .NET Framework 4.8.1 Developer Pack
4. Microsoft Excel installed (for highest fidelity primary converter)

## NuGet packages
- Microsoft.Office.Interop.Excel
- Microsoft.AspNet.WebApi
- Microsoft.AspNet.WebApi.Core
- Microsoft.AspNet.WebApi.WebHost
- ClosedXML
- PdfSharp

## Build and run (Visual Studio)
1. Open `SpreadsheetToPdf.sln`.
2. Restore NuGet packages.
3. Build solution.
4. Run with IIS Express or local IIS.
5. Call `GET /api/conversion/health` to verify service startup.

## Why Excel Interop remains preferred
Excel Interop uses Excel's native rendering and print engine, so merged cells, formatting, pagination, print areas, and sheet layout are preserved much more accurately than generic library-only rendering.

## Visual Studio startup fix
If Visual Studio shows **"A project with an output type of class library cannot be started directly"**:
1. Ensure the solution project type is **ASP.NET Web Application** (already configured in this repo).
2. Right-click `SpreadsheetToPdf` project → **Set as Startup Project**.
3. In project properties, go to **Web** tab and use **IIS Express** as the Start Action.
4. Press F5 again.

Note: Web API projects intentionally use `OutputType=Library`; they are hosted by IIS/IIS Express, not started like console executables.
