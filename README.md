# ConvertToPdf - ASP.NET Web API Self-Host (.NET Framework 4.8.1)

This project runs as a **self-hosted ASP.NET Web API executable** on Windows (.NET Framework 4.8.1), so Visual Studio can start it directly (no IIS project flavor required).

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
curl -X POST "http://localhost:5000/api/conversion/pdf" \
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
- Microsoft.AspNet.WebApi.Core
- Microsoft.AspNet.WebApi.SelfHost
- ClosedXML
- PdfSharp

## Build and run (Visual Studio)
1. Open `SpreadsheetToPdf.sln`.
2. Restore NuGet packages.
3. Build solution.
4. Run (`F5`) — this starts the self-host executable.
5. Call `GET http://localhost:5000/api/conversion/health`.

## Configuring host URL
Set `BaseAddress` in `App.config` (default: `http://localhost:5000`).

## Fix for "A project with an output type of class library cannot be started directly"
This repository now uses `OutputType=Exe` with Web API self-host, so Visual Studio can start it directly.

## Why Excel Interop remains preferred
Excel Interop uses Excel's native rendering and print engine, so merged cells, formatting, pagination, print areas, and sheet layout are preserved much more accurately than generic library-only rendering.
