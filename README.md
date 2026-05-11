# ExcelIO

ExcelIO is a lightweight, high-performance C# library for reading and writing Excel files. It supports both modern OpenXML (.xlsx, .xlsm) and legacy binary (.xls) formats.

## Features

- **XLSX & XLS Support**: Read and write modern and legacy formats.
- **Fast Performance**: Optimized for speed and low memory footprint.
- **High-Level API**: Easy-to-use object model (`Workbook`, `Worksheet`, `Row`, `Cell`).
- **Rich Styling**: (New) Support for fonts, fills, alignments, borders, column widths, row heights, and sheet options.
- **Image Support**: Insert images into worksheets with precise positioning.
- **Zero Dependencies**: Core library does not depend on heavy external packages like OpenXML SDK or Interop.
- **Async API**: Support for asynchronous load and save operations.

## Installation

```bash
dotnet add package ExcelIO
```

## Quick Start

### Basic Writing

```csharp
using ExcelIO;

var wb = new XlWorkbook();
var ws = wb.NewWorksheet("Sheet1");

ws.AddRow("Name", "Age", "City");
ws.AddRow("Alice", "25", "New York");
ws.AddRow("Bob", "30", "London");

XlHelper.Save("output.xlsx", wb);
```

### Advanced Styling

```csharp
var wb = new XlWorkbook();
var ws = wb.NewWorksheet("StyledSheet");

// Sheet Options
ws.Options.TabColor = "FFFF0000"; // ARGB Hex
ws.Options.ShowGridLines = false;
ws.Options.DefaultRowHeight = 20;

// Column Properties
ws.Columns[0] = new XlColumn { Width = 30 };

// Row and Cell Styling
var row = ws.NewRow(["Header 1", "Header 2"]);
row.Height = 30;
row.Style = new XlStyle 
{ 
    Bold = true, 
    FontSize = 12, 
    FillColor = "FFD3D3D3", // Light Gray
    Alignment = new XlAlignment { Horizontal = XlHorizontalAlignment.Center }
};

// Specific Cell Override
row.Cells[0].Style = new XlStyle { FontColor = "FFFF0000" }; // Red Text

XlHelper.Save("styled.xlsx", wb);
```

### Reading Excel Files

```csharp
var workbook = XlHelper.Load("input.xlsx");
var sheet = workbook.Worksheets[0];

foreach (var row in sheet)
{
    foreach (var cell in row.Cells)
    {
        Console.Write(cell.Value + "\t");
    }
    Console.WriteLine();
}
```

### Image Support

```csharp
var ws = wb.NewWorksheet("Images");
ws.AddImage("logo.png", rowIndex: 1, columnIndex: 1, rowSpan: 5, columnSpan: 3);
// NEW! Support place images into cell
ws.AddImage("logo.png", rowIndex: 1, columnIndex: 1, rowSpan: 5, columnSpan: 3, placeInCell: true);

XlHelper.Save("with_images.xlsx", wb);

```

## License

This project is licensed under the MIT License.
