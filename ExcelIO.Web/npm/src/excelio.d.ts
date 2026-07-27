export interface InitConfig {
  /** Base path for `_framework/` WASM assets. Default: `'.'` */
  baseUrl?: string;
  /** Enable WASM diagnostic tracing. Default: false */
  diagnosticTracing?: boolean;
}

export interface CellStyle {
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  fillColor?: string;
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

export interface ExcelIO {
  /** Load an Excel file from bytes. Returns true on success. */
  loadFromBytes(data: Uint8Array, fileName: string): boolean;

  /** Serialize the current workbook to .xlsx bytes. */
  saveToBytes(): Uint8Array;

  /** Create an empty workbook with one sheet. */
  newWorkbook(): void;

  /** Number of worksheets in the workbook. */
  getSheetCount(): number;

  /** Sheet name by index (0-based). */
  getSheetName(index: number): string;

  /** Active sheet index (0-based). */
  getActiveSheetIndex(): number;

  /** Switch active sheet by index. Returns false if index is out of range. */
  switchSheet(index: number): boolean;

  /** Current filename (without path). */
  getFileName(): string;

  /** Row count in the active sheet. */
  getRowCount(): number;

  /** Maximum column count in the active sheet. */
  getColCount(): number;

  /** Cell text at (row, col). Returns empty string if out of range. */
  getCellValue(row: number, col: number): string;

  /** Cell formula at (row, col), e.g. "=SUM(A1:A2)". Empty string if not a formula. */
  getCellFormula(row: number, col: number): string;

  /** Set cell text. Creates the row and cell if they don't exist. */
  setCellValue(row: number, col: number, value: string): void;

  /** Append a row of values. */
  addRow(values: string[]): void;

  /** Remove all rows from the active sheet. */
  clearRows(): void;

  /** Cell-level style as JSON. Resolves cascade: cell > row > column. */
  getCellStyleJson(row: number, col: number): string;

  /** Row-level style as JSON. */
  getRowStyleJson(row: number): string;

  /** Column-level style as JSON. */
  getColStyleJson(col: number): string;

  /** Merged cell ranges as a JSON array of strings (e.g. `["A1:B2"]`). */
  getMergedCells(): string;

  /** Column width in points, or 0 if default. */
  getColumnWidth(col: number): number;

  /** Row height in points, or 0 if default. */
  getRowHeight(row: number): number;

  /** Whether the row is hidden. */
  isRowHidden(row: number): boolean;

  /** Whether the column is hidden. */
  isColumnHidden(col: number): boolean;

  /** Active sheet tab color as CSS hex (#RRGGBB), or empty string. */
  getSheetTabColor(): string;
}

/**
 * Initialize the ExcelIO WebAssembly runtime.
 *
 * Singleton — subsequent calls return the cached promise.
 *
 * @example
 * import init from 'excelio-web';
 * const excelIO = await init();
 * excelIO.newWorkbook();
 */
export default function init(config?: InitConfig): Promise<ExcelIO>;
