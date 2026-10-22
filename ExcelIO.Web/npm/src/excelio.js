let _ready = null;

/**
 * Initialize the ExcelIO WebAssembly runtime.
 * Returns a singleton — subsequent calls return the same promise.
 *
 * @param {Object} [config]
 * @param {string} [config.baseUrl] — base path for `_framework/` assets (default: `'.'`)
 * @param {boolean} [config.diagnosticTracing] — enable WASM diagnostic tracing (default: false)
 * @returns {Promise<ExcelIO>}
 */
function init(config = {}) {
  if (_ready) return _ready;

  _ready = (async () => {
    const base = config.baseUrl || '.';
    const { dotnet } = await import(`${base}/_framework/dotnet.js`);

    const { getAssemblyExports, getConfig } = await dotnet
      .withDiagnosticTracing(!!config.diagnosticTracing)
      .create();

    const bootConfig = getConfig();
    const exports = await getAssemblyExports(bootConfig.mainAssemblyName);
    const api = exports.ExcelIO.ExcelIOWrapper;

    const excelIO = {
      // ── Workbook ──
      loadFromBytes(data, fileName) {
        return api.LoadFromBytes(data, fileName);
      },
      saveToBytes() {
        return api.SaveToBytes();
      },
      newWorkbook() {
        api.NewWorkbook();
      },
      getSheetCount() {
        return api.GetSheetCount();
      },
      getSheetName(index) {
        return api.GetSheetName(index);
      },
      getActiveSheetIndex() {
        return api.GetActiveSheetIndex();
      },
      switchSheet(index) {
        return api.SwitchSheet(index);
      },
      getFileName() {
        return api.GetFileName();
      },

      // ── Data ──
      getRowCount() {
        return api.GetRowCount();
      },
      getColCount() {
        return api.GetColCount();
      },
      getCellValue(row, col) {
        return api.GetCellValue(row, col);
      },
      getCellFormula(row, col) {
        return api.GetCellFormula(row, col);
      },
      setCellValue(row, col, value) {
        // WASM interop requires explicitly passing string for the value param
        api.SetCellValue(row, col, String(value));
      },
      addRow(values) {
        api.AddRow(values);
      },
      clearRows() {
        api.ClearRows();
      },

      // ── Style ──
      getCellStyleJson(row, col) {
        return api.GetCellStyleJson(row, col);
      },
      getRowStyleJson(row) {
        return api.GetRowStyleJson(row);
      },
      getColStyleJson(col) {
        return api.GetColStyleJson(col);
      },
      getMergedCells() {
        return api.GetMergedCells();
      },
      getColumnWidth(col) {
        return api.GetColumnWidth(col);
      },
      getRowHeight(row) {
        return api.GetRowHeight(row);
      },
      isRowHidden(row) {
        return api.IsRowHidden(row);
      },
      isColumnHidden(col) {
        return api.IsColumnHidden(col);
      },
      getSheetTabColor() {
        return api.GetSheetTabColor();
      },

      // ── Formula ──
      setCellFormula(row, col, formula) {
        api.SetCellFormula(row, col, String(formula));
      },
      calculateSheet() {
        api.CalculateSheet();
      },
      calculateWorkbook() {
        api.CalculateWorkbook();
      },
      getCircularReferences() {
        return api.GetCircularReferences();
      },

      // ── Style setters ──
      setCellStyle(row, col, styleJson) {
        api.SetCellStyle(row, col, String(styleJson));
      },
      setRowStyle(row, styleJson) {
        api.SetRowStyle(row, String(styleJson));
      },
      setColStyle(col, styleJson) {
        api.SetColStyle(col, String(styleJson));
      },

      // ── Layout setters ──
      setColumnWidth(col, width) {
        api.SetColumnWidth(col, width);
      },
      setRowHeight(row, height) {
        api.SetRowHeight(row, height);
      },

      // ── Merge ──
      mergeCells(rangeRef) {
        api.MergeCells(String(rangeRef));
      },
      unmergeCells(rangeRef) {
        api.UnmergeCells(String(rangeRef));
      },

      // ── Sheet management ──
      addSheet(name) {
        return api.AddSheet(String(name));
      },
      deleteSheet(index) {
        return api.DeleteSheet(index);
      },
      renameSheet(index, newName) {
        return api.RenameSheet(index, String(newName));
      },
    };

    await dotnet.run();
    return excelIO;
  })();

  return _ready;
}

export default init;
