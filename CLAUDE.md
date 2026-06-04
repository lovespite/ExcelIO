# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and test commands

```powershell
# Build library
dotnet build .\ExcelIO\ExcelIO.csproj -v minimal

# Run all tests
dotnet test .\ExcelIO.Test\ExcelIO.Test.csproj -v minimal

# Run a single test method
dotnet test .\ExcelIO.Test\ExcelIO.Test.csproj --filter "FullyQualifiedName~TestClassName.TestMethodName" -v minimal
```

## Architecture

ExcelIO is a zero-dependency C# library (targeting `net10.0`) for reading/writing Excel files in both modern OpenXML (`.xlsx`, `.xlsm`) and legacy binary (`.xls`) formats. All XML generation and parsing is manual -- no OpenXML SDK or Interop dependency.

### Public model (`ExcelIO` namespace)

- **`XlWorkbook`** → **`XlWorksheet`** → **`XlRow`** → **`XlCell`** — the in-memory hierarchy shared by all formats. Each is enumerable (`IReadOnlyList<T>`) with indexers.
- **`XlHelper`** — static facade for `Load()` and `Save()`. All format dispatch lives here.
- **`XlRange`** — range-based operations on a worksheet (set content, set/clear style, merge/unmerge). Supports A1, A, A:C, 1:3, and A1:C3 notation. Infinite-row and infinite-column ranges are supported (e.g., "A" means all rows in column A) but silently no-op on `SetContent()`.
- **`XlStyle`** / `XlAlignment` / `XlBorder` — styling model. Colors are ARGB hex strings (`"FFFF0000"`). Styles can be applied at the column, row, or cell level.
- **`XlLoadOptions`** — controls whether styles/images are loaded on read (both default `true`).
- **`XlSheetOptions`** — per-sheet properties: `TabColor`, `ShowGridLines`, `DefaultRowHeight`.
- **`XlWorksheetImage`** — image metadata and bytes. Supports PNG, JPG, GIF, BMP, TIFF.

### OpenXML write path (`XlHelper.Save`)

Manually constructs a ZIP archive with the required OPC parts: `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/styles.xml`, `xl/worksheets/sheetN.xml`, plus optional drawing/rich-data parts for images.

- **Cell values use `inlineStr`** — no `sharedStrings.xml` is written. All cell text is serialized inline.
- **`XlsxStyleBuilder`** collects unique `XlStyle` instances during generation and emits `xl/styles.xml`.
- **Floating images** use `twoCellAnchor` drawings with relationships in `xl/drawings/`.
- **In-cell images** (`placeInCell: true`) use the modern rich-data stack: `xl/metadata.xml`, `xl/richData/rdrichvalue.xml`, `xl/richData/rdrichvaluestructure.xml`, etc.

### OpenXML read path (`XlHelper.LoadOpenXml`)

Parses the ZIP via `System.IO.Compression.ZipArchive` and `XDocument`/`XmlReader`:
1. Reads `xl/sharedStrings.xml` into a list.
2. Reads `xl/styles.xml` via `XlsxStyleReader`.
3. Reads `xl/workbook.xml` for sheet names and `r:id` references.
4. Resolves relationships in `xl/_rels/workbook.xml.rels`.
5. Parses each worksheet XML with `XmlReader` for rows, cells, styles, column definitions, sheet options, and floating images.

### Legacy `.xls` read path

- **`XlsCompoundReader`** — extracts the "Workbook" or "Book" stream from the OLE2 Compound File Binary (CFB) container. Handles FAT, DIFAT, MiniFAT chain resolution.
- **`XlsBiff8Reader`** — parses BIFF8 records from the workbook stream into the in-memory model. Currently supports: `LABELSST`, `NUMBER`, `RK`, `LABEL`, `SST`, `CONTINUE`, `BOUNDSHEET`.

## Key conventions

- `XlCell.Null` is a null-object sentinel. Check `cell.IsNull` before accessing its `Value`.
- `XlRow[int]` (indexer) returns `string.Empty` for out-of-range column indexes — it does not throw.
- `XlWorksheet.MapHeaders(...)` caches column-name-to-index mappings in a case-insensitive dictionary. It may create a first row when the sheet is empty.
- New `.xls` parsing tests construct minimal BIFF8/CFB payloads inline in C#; do not rely on external fixture `.xls` files.
- Format-specific parsing stays in dedicated helpers (`XlsCompoundReader`, `XlsBiff8Reader`). Format dispatch stays in `XlHelper`.
- The project was formerly under `ClassLibrary1\` but now lives at `ExcelIO\`. The solution file is `ExcelIO.slnx` (the new XML-based format).
