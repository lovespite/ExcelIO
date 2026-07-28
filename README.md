# ExcelIO

ExcelIO is a lightweight, high-performance C# library for reading and writing Excel files. It supports both modern OpenXML (.xlsx, .xlsm) and legacy binary (.xls) formats.

## Features

- **XLSX & XLS Support**: Read and write modern and legacy formats.
- **Formula Engine** (`ExcelIO.Formula`): Pluggable formula calculation with 25+ built-in functions, custom function registration, dependency tracking, and circular reference detection.
- **BIFF8 Formula Support**: Decompile legacy .xls formula bytecode into evaluable formula strings.
- **Fast Performance**: Optimized for speed and low memory footprint.
- **High-Level API**: Easy-to-use object model (`Workbook`, `Worksheet`, `Row`, `Cell`).
- **Rich Styling**: Support for fonts, fills, alignments, borders, column widths, row heights, and sheet options.
- **Image Support**: Insert images into worksheets with precise positioning (floating or in-cell).
- **Zero Dependencies**: Core library does not depend on heavy external packages like OpenXML SDK or Interop.
- **Async API**: Support for asynchronous load and save operations.

## Installation

```bash
dotnet add package ExcelIO
```

For formula calculation support, also install the formula engine:

```bash
dotnet add package ExcelIO.Formula
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

### Formula Calculation

```csharp
using ExcelIO;
using ExcelIO.Formula;

// Enable the formula engine
XlHelper.FormulaEngine = new FormulaEngine();

var wb = new XlWorkbook();
var ws = wb.NewWorksheet("Sheet1");

// Set up data
ws.NewRow(["10", "20", "30"]);

// Add formulas
var row = ws.NewRow();
var sumCell = row.Insert(0, "");
sumCell.SetFormula("=SUM(A1:C1)");          // = 60
var avgCell = row.Insert(1, "");
avgCell.SetFormula("=AVERAGE(A1:C1)");       // = 20
var ifCell = row.Insert(2, "");
ifCell.SetFormula("=IF(A1>5,\"big\",\"small\")"); // = "big"

// Calculate all formulas (topological order)
XlHelper.FormulaEngine.Calculate(ws);

Console.WriteLine(sumCell.Value);  // "60"
Console.WriteLine(avgCell.Value);  // "20"
Console.WriteLine(ifCell.Value);   // "big"
```

### Custom Functions

Register your own functions using the `ExcelFunction` descriptor, `[ExcelFunction]` attribute, or assembly scanning:

```csharp
var engine = new FormulaEngine();

// Style 1 — inline registration
engine.Functions.Register(new ExcelFunction(
    "DOUBLE", "Custom", "Multiplies by 2",
    minArgs: 1, maxArgs: 1,
    (args, ctx) => Convert.ToDouble(args[0]) * 2));

// Style 2 — class-based registration with attributes
public class FinancialFunctions : IFormulaFunction
{
    [ExcelFunction(Name = "PV", Category = "Financial",
        Description = "Present value of an investment",
        MinArgs = 3, MaxArgs = 3)]
    public static object Pv(IReadOnlyList<object> args, IFormulaContext ctx)
    {
        var rate = Convert.ToDouble(args[0]);
        var nper = Convert.ToDouble(args[1]);
        var pmt  = Convert.ToDouble(args[2]);
        return pmt * (1 - Math.Pow(1 + rate, -nper)) / rate;
    }
}
engine.Functions.RegisterAll<FinancialFunctions>();

// Style 3 — scan an entire assembly
engine.Functions.RegisterAssembly(typeof(MyAddin).Assembly);

// Discover available functions
foreach (var fn in engine.Functions.GetByCategory("Financial"))
    Console.WriteLine($"{fn.Name}: {fn.Description}");
```

### Built-in Functions

| Category | Functions |
|---|---|
| Math | `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ROUND`, `ABS`, `PRODUCT` |
| Logic | `IF`, `AND`, `OR`, `NOT`, `IFERROR` |
| Lookup | `VLOOKUP`, `MATCH`, `INDEX` |
| Text | `CONCATENATE`, `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `UPPER`, `LOWER` |
| Date | `TODAY`, `NOW` |

### Dependency Tracking & Circular Reference Detection

The engine builds a dependency graph and evaluates cells in topological order. Circular references are detected and reported:

```csharp
var engine = new FormulaEngine();
XlHelper.FormulaEngine = engine;

// A1 = B1 + 1, B1 = A1 + 1  → circular
a1.SetFormula("=B1+1");
b1.SetFormula("=A1+1");

engine.Calculate(ws);

if (engine.CircularReferences.Count > 0)
{
    foreach (var circ in engine.CircularReferences)
        Console.WriteLine($"Circular: {string.Join(" → ", circ.Path)}");
}
```

### Reading Formulas from .xls Files

When `ExcelIO.Formula` is referenced, BIFF8 formula bytecode in legacy .xls files is automatically decompiled into A1-notation formula strings:

```csharp
using ExcelIO.Formula;

// Install the engine (auto-wires the BIFF8 decompiler)
XlHelper.FormulaEngine = new FormulaEngine();

// Load an .xls with formulas — they're decompiled and evaluable
var wb = XlHelper.Load("legacy.xls");
var cell = wb.Worksheets[0].Rows[0].Cells[0];

Console.WriteLine(cell.Formula); // e.g. "=SUM(A1:A10)"
Console.WriteLine(cell.Value);   // cached result from the file
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
// In-cell images (embedded in the cell rather than floating)
ws.AddImage("logo.png", rowIndex: 1, columnIndex: 1, rowSpan: 5, columnSpan: 3, placeInCell: true);

XlHelper.Save("with_images.xlsx", wb);
```

## Architecture

```
ExcelIO                 → Core library (net10.0, zero-dependency)
  ├── XlWorkbook / XlWorksheet / XlRow / XlCell   — in-memory model
  ├── XlHelper.Load() / Save()                    — format dispatch
  ├── XlRange                                      — A1-notation range operations
  ├── XlStyle / XlAlignment / XlBorder            — styling model
  ├── XlsCompoundReader / XlsBiff8Reader          — legacy .xls read
  └── XlSharedFormulaOptimizer                    — write-path shared formula

ExcelIO.Formula        → Formula engine (net10.0, references ExcelIO)
  ├── FormulaEngine / FormulaEvaluator            — calculation engine
  ├── FormulaParser / Ast                          — recursive-descent parser
  ├── FunctionRegistry / BuiltinFunctions          — extensible function system
  ├── DependencyGraph / CircularReferenceDetector  — recalculation order
  └── Biff8FormulaReader                           — BIFF8 RPN decompiler

ExcelIO.NetStandard     → netstandard2.0 compatibility wrapper
ExcelIO.Web             → Blazor WebAssembly wrapper
ExcelIO.Benchmark       → BenchmarkDotNet performance tests
```

## License

This project is licensed under the MIT License.
