# ConvertToPdf

## SpreadsheetToPdf (.NET Framework 4.8.1)

A production-oriented Windows (.NET Framework 4.8.1) console application that converts spreadsheet files to PDF using **Microsoft Excel Interop** for maximum layout fidelity.

### Features
- Converts `.xlsx`, `.xls`, and `.csv` to `.pdf`.
- Uses Excel's built-in `ExportAsFixedFormat` engine for high-accuracy output.
- Preserves workbook rendering details as Excel prints them (including merged cells, formatting, borders, fonts, colors, and print layout).
- Supports exporting:
  - all worksheets, or
  - a specific worksheet by name.
- Detects CSV delimiter from sampled content and imports with Excel before exporting.
- Handles robust error scenarios with user-friendly messages and explicit exit codes.
- Uses defensive COM cleanup to reduce orphaned `Excel.exe` processes.

## Prerequisites
1. **Windows OS**.
2. **Microsoft Excel installed** (desktop version with COM automation support).
3. **Visual Studio 2022** (or compatible version that supports .NET Framework 4.8.1).
4. **.NET Framework Developer Pack 4.8.1** installed.

## Excel Interop reference / NuGet setup
This project references:
- `Microsoft.Office.Interop.Excel` (NuGet package)

If you need to add it manually in Visual Studio:
1. Right-click project → **Manage NuGet Packages**.
2. Search for `Microsoft.Office.Interop.Excel`.
3. Install the package into the `SpreadsheetToPdf` project.

Alternatively using Package Manager Console:
```powershell
Install-Package Microsoft.Office.Interop.Excel -Version 15.0.4795.1001
```

## Build in Visual Studio
1. Open `SpreadsheetToPdf.sln`.
2. Ensure configuration is set to **Release** (or Debug).
3. Build solution (**Build → Build Solution**).
4. Output executable is generated under:
   - `SpreadsheetToPdf\bin\Release\SpreadsheetToPdf.exe`

## Command-line usage
```powershell
SpreadsheetToPdf.exe "C:\Files\report.xlsx" "C:\Files\report.pdf"
SpreadsheetToPdf.exe "C:\Files\report.xls" "C:\Files\report.pdf" "Sheet1"
SpreadsheetToPdf.exe "C:\Files\data.csv" "C:\Files\data.pdf"
```

### Arguments
1. `input file path` (`.xlsx`, `.xls`, `.csv`)
2. `output pdf path` (`.pdf`)
3. optional worksheet name (if provided, only that worksheet is exported)

## Exit codes
- `0` success
- `1` invalid arguments
- `2` input file not found
- `3` invalid / unsupported format
- `4` Excel not installed / cannot start Excel Interop
- `5` worksheet not found
- `6` access denied
- `7` Excel COM/interoperability failure
- `99` unexpected failure

## Limitations of relying on Microsoft Excel
- Requires Excel to be installed on the machine running the app.
- Depends on COM automation, so server-side unattended usage needs careful operational controls.
- Output fidelity can vary slightly by Excel version, installed fonts, printer defaults, and regional settings.
- Not cross-platform; this implementation is Windows-only by design.

## Why Excel Interop is best for highest spreadsheet-to-PDF accuracy
Excel Interop uses Excel's own rendering and print/export engine. That means PDF output follows how Excel itself interprets workbook layout, formatting, print areas, merged regions, pagination, fonts, and styling. Generic spreadsheet libraries can be excellent for data processing, but they often reimplement rendering behavior and may not match Excel's print fidelity for complex real-world workbooks.
