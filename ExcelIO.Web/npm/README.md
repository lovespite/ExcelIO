# excelio-web

ExcelIO WebAssembly — read and write Excel files (`.xlsx`, `.xlsm`, `.xls`) in the browser.

Built on [ExcelIO](https://github.com/lovespite/ExcelIO), a zero-dependency C# library compiled to WebAssembly via .NET AOT.

## Install

```bash
npm install excelio-web
```

## Quick Start

```js
import init from 'excelio-web';

const excelIO = await init();

// Create a new workbook
excelIO.newWorkbook();
excelIO.setCellValue(0, 0, 'Hello');
excelIO.setCellValue(0, 1, 'World');

// Save to bytes and trigger download
const bytes = excelIO.saveToBytes();
const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'output.xlsx';
a.click();
URL.revokeObjectURL(url);
```

## Load an existing file

```js
const input = document.querySelector('input[type="file"]');
input.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const buffer = await file.arrayBuffer();
  const data = new Uint8Array(buffer);
  excelIO.loadFromBytes(data, file.name);
  // Now read data...
  console.log(excelIO.getCellValue(0, 0));
});
```

## API

### Workbook

| Method | Description |
|--------|-------------|
| `loadFromBytes(data, fileName)` | Load Excel from `Uint8Array`. Returns `boolean`. |
| `saveToBytes()` | Serialize workbook to `Uint8Array` (xlsx format). |
| `newWorkbook()` | Create an empty workbook. |
| `getSheetCount()` | Number of worksheets. |
| `getSheetName(index)` | Sheet name by index. |
| `getActiveSheetIndex()` | Current sheet index (0-based). |
| `switchSheet(index)` | Switch active sheet. Returns `boolean`. |
| `getFileName()` | Current filename without path. |

### Data

| Method | Description |
|--------|-------------|
| `getRowCount()` | Row count in active sheet. |
| `getColCount()` | Max column count in active sheet. |
| `getCellValue(row, col)` | Cell text at (row, col). |
| `setCellValue(row, col, value)` | Set cell text. Creates row/cell if needed. |
| `addRow(values)` | Append a row from `string[]`. |
| `clearRows()` | Remove all rows from active sheet. |

### Style

Methods returning JSON strings — use `JSON.parse()` to get typed objects:

| Method | Description |
|--------|-------------|
| `getCellStyleJson(row, col)` | Cell style (cascade: cell > row > column). |
| `getRowStyleJson(row)` | Row-level style. |
| `getColStyleJson(col)` | Column-level style. |
| `getMergedCells()` | JSON array of merged cell ranges (e.g. `["A1:B2"]`). |
| `getColumnWidth(col)` | Column width in points, or `0` if default. |
| `getRowHeight(row)` | Row height in points, or `0` if default. |
| `isRowHidden(row)` | Whether the row is hidden. |
| `isColumnHidden(col)` | Whether the column is hidden. |
| `getSheetTabColor()` | Tab color as CSS hex, or `""`. |

### CellStyle shape

```ts
interface CellStyle {
  fontName?: string;
  fontSize?: number;
  fontColor?: string;     // CSS hex
  bold?: boolean;
  italic?: boolean;
  fillColor?: string;     // CSS hex
  hAlign?: 'left' | 'center' | 'right' | 'justify';
  vAlign?: 'top' | 'middle' | 'bottom' | 'justify';
  wrapText?: boolean;
  borderLeft?: string;
  borderLeftColor?: string;
  borderRight?: string;
  borderRightColor?: string;
  borderTop?: string;
  borderTopColor?: string;
  borderBottom?: string;
  borderBottomColor?: string;
}
```

## CDN hosting

Host `_framework/` on a CDN and pass `baseUrl`:

```js
const excelIO = await init({ baseUrl: 'https://cdn.example.com/excelio' });
```

## Bundler setup

The WASM runtime resolves assets (`_framework/`) relative to `excelio.js`. Your bundler must copy `node_modules/excelio-web/dist/_framework/` to your output directory.

**Vite** — use `vite-plugin-static-copy`:

```js
import { viteStaticCopy } from 'vite-plugin-static-copy';
export default {
  plugins: [
    viteStaticCopy({
      targets: [{
        src: 'node_modules/excelio-web/dist/_framework/*',
        dest: '_framework'
      }]
    })
  ]
};
```

## License

MIT
