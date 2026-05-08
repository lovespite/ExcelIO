# Copilot instructions for this repository

## Build and test commands

Use project-file commands directly (the `ExcelIO.slnx` file currently references a different project path).

```powershell
# Build library
dotnet build .\ClassLibrary1\ExcelIO.csproj -v minimal

# Run all tests
dotnet test .\ExcelIO.Test\ExcelIO.Test.csproj -v minimal

# Run a single test method
dotnet test .\ExcelIO.Test\ExcelIO.Test.csproj --filter "FullyQualifiedName~XlHelperXlsLoadTests.Load_Xls_ReadsMinimalTextAndNumberCells" -v minimal
```

## High-level architecture

- `ClassLibrary1\XlHelper.cs` is the public facade for load/save.
  - `Save(...)` writes **OpenXML `.xlsx`** by manually generating ZIP entries/XML.
  - `Load(string path)` routes by file extension (`.xls` vs `.xlsx/.xlsm`).
  - `Load(Stream)` / `Load(ReadOnlySpan<byte>)` route by file signature (ZIP vs CFB header).
- OpenXML read path (`LoadOpenXml`) manually parses:
  - `xl/sharedStrings.xml`
  - `xl/workbook.xml`
  - `xl/_rels/workbook.xml.rels`
  - each worksheet XML stream
- Legacy `.xls` read path is split:
  - `XlsCompoundReader` extracts the `Workbook`/`Book` stream from CFB.
  - `XlsBiff8Reader` parses BIFF8 records into the in-memory model (currently focused on text/number-related records).
- In-memory model hierarchy:
  - `XlWorkbook` -> `XlWorksheet` -> `XlRow` -> `XlCell`
  - these types are collection-like (`IReadOnlyList` implementations + indexers) and are the shared contract for all formats.

## Key conventions in this codebase

- Keep format handling centralized in `XlHelper`; format-specific parsing stays in dedicated helpers (`XlsCompoundReader`, `XlsBiff8Reader`).
- The `.xlsx` writer uses `inlineStr` cells and does **not** emit `sharedStrings.xml`; cell values are serialized as strings.
- `XlWorksheet` header-based access relies on `MapHeaders(...)` and a case-insensitive `_indexCache`.
  - `MapHeaders` may create the first row when the sheet is empty.
- Null/out-of-range behavior is intentionally non-throwing in several accessors:
  - `XlCell.Null` is used as a null-object sentinel.
  - `XlRow[int]` returns `string.Empty` for out-of-range indexes.
- Tests in `ExcelIO.Test\XlHelperXlsLoadTests.cs` construct minimal BIFF8/CFB payloads directly in code; keep new `.xls` parsing tests close to this pattern rather than relying on external fixture files.
