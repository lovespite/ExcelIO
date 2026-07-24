import { dotnet } from './_framework/dotnet.js'

const { getAssemblyExports, getConfig } = await dotnet
    .withDiagnosticTracing(false)
    .withApplicationArgumentsFromQuery()
    .create();

const config = getConfig();
const exports = await getAssemblyExports(config.mainAssemblyName);
const api = exports.ExcelIO.ExcelIOWrapper;

window.excelIO = {
    // Workbook
    loadFromBytes:    api.LoadFromBytes,
    saveToBytes:      api.SaveToBytes,
    newWorkbook:      api.NewWorkbook,
    getSheetCount:    api.GetSheetCount,
    getSheetName:     api.GetSheetName,
    getActiveSheetIndex: api.GetActiveSheetIndex,
    switchSheet:      api.SwitchSheet,
    getFileName:      api.GetFileName,

    // Data
    getRowCount:      api.GetRowCount,
    getColCount:      api.GetColCount,
    getCellValue:     api.GetCellValue,
    setCellValue:     api.SetCellValue,
    addRow:           api.AddRow,
    clearRows:        api.ClearRows,

    // Style
    getCellStyleJson: api.GetCellStyleJson,
    getRowStyleJson:  api.GetRowStyleJson,
    getColStyleJson:  api.GetColStyleJson,
    getMergedCells:   api.GetMergedCells,
    getColumnWidth:   api.GetColumnWidth,
    getRowHeight:     api.GetRowHeight,
    isRowHidden:      api.IsRowHidden,
    isColumnHidden:   api.IsColumnHidden,
    getSheetTabColor: api.GetSheetTabColor,
};

await dotnet.run();
