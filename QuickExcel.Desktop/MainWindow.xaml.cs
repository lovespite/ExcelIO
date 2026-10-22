using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ExcelIO;
using Microsoft.Win32;
using QuickExcel.Core;

namespace QuickExcel.Desktop;

public partial class MainWindow : Window
{
    private WorkbookDocument _document = WorkbookDocument.CreateDemo();
    private readonly ObservableCollection<SheetRow> _rows = [];
    private readonly Stack<EditBatch> _undo = new();
    private readonly Stack<EditBatch> _redo = new();
    private int _sheetIndex;
    private int _rowCount = 100;
    private int _columnCount = 12;
    private int _selectedRow;
    private int _selectedColumn;
    private bool _refreshing;
    private bool _busy;
    private bool _dirty;
    private bool _closingApproved;
    private bool _calculationPending;
    private bool _suppressAutoCalculation;
    private string? _filePath;
    private string _displayName = "示例工作簿.xlsq";

    public MainWindow()
    {
        InitializeComponent();
        SheetGrid.ItemsSource = _rows;
        InstallDocument(_document, null, "示例工作簿.xlsq", disposePrevious: false);
    }

    public async Task StartAsync(string[] args)
    {
        if (args.Length > 0 && args[0] == "--smoke-test")
        {
            try
            {
                await RunSmokeTestAsync(args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "smoke"));
                Environment.ExitCode = 0;
            }
            catch (Exception ex)
            {
                Environment.ExitCode = 1;
                var folder = args.Length > 1 ? args[1] : AppContext.BaseDirectory;
                Directory.CreateDirectory(folder);
                File.WriteAllText(Path.Combine(folder, "smoke-error.txt"), ex.ToString());
            }
            _closingApproved = true;
            Close();
            return;
        }
        if (args.Length > 0 && File.Exists(args[0])) await OpenPathAsync(args[0]);
    }

    private void InstallDocument(WorkbookDocument document, string? path, string displayName, bool disposePrevious = true)
    {
        if (disposePrevious) _document.Dispose();
        _document = document;
        _filePath = path != null && Path.GetExtension(path).Equals(".xlsq", StringComparison.OrdinalIgnoreCase) ? path : null;
        _displayName = displayName;
        _sheetIndex = 0;
        _selectedRow = _selectedColumn = 0;
        _rowCount = 100;
        _columnCount = 12;
        _undo.Clear();
        _redo.Clear();
        _dirty = false;
        _calculationPending = false;
        _refreshing = true;
        ScriptEditor.Text = document.ScriptSource;
        SheetPicker.ItemsSource = document.Workbook.Worksheets.Select(s => s.Name).ToList();
        SheetPicker.SelectedIndex = 0;
        _refreshing = false;
        RefreshSheet(true);
        UpdateChrome();
        ScriptState.Text = "● 脚本未启用";
        LogText.Text = "工作簿已载入。\n点击「应用脚本并重算」启用此工作簿的 JavaScript。";
        HintText.Text = document.Workbook.Worksheets.Any(s => s.Rows.Any(r => r.Cells.Count > 256))
            ? "此工作簿超过界面的 256 列显示范围；未显示的受支持数据仍会保存。脚本需手动启用。"
            : "开始试用：点击右侧「应用脚本并重算」，再修改价格或数量。双击单元格编辑；公式栏可查看公式。";
    }

    private void RefreshSheet(bool rebuild = false)
    {
        var sheet = _document.Workbook.Worksheets[_sheetIndex];
        if (rebuild)
        {
            _rowCount = Math.Clamp(Math.Max(_rowCount, sheet.Rows.Count + 20), 100, 100000);
            _columnCount = Math.Clamp(Math.Max(_columnCount, sheet.Rows.Select(r => r.Cells.Count).DefaultIfEmpty().Max() + 2), 12, 256);
            SheetGrid.Columns.Clear();
            for (var c = 0; c < _columnCount; c++)
            {
                var textStyle = new Style(typeof(TextBlock));
                textStyle.Setters.Add(new Setter(TextBlock.PaddingProperty, new Thickness(8, 4, 5, 3)));
                textStyle.Setters.Add(new Setter(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis));
                textStyle.Setters.Add(new Setter(TextBlock.ForegroundProperty, new Binding { Converter = new AppearanceConverter(c, "foreground") }));
                textStyle.Setters.Add(new Setter(TextBlock.BackgroundProperty, new Binding { Converter = new AppearanceConverter(c, "background") }));
                textStyle.Setters.Add(new Setter(TextBlock.FontWeightProperty, new Binding { Converter = new AppearanceConverter(c, "weight") }));
                var width = c == 0 ? 160d : 112d;
                if (sheet.Columns.TryGetValue(c, out var column) && column.Width.HasValue)
                    width = Math.Clamp(column.Width.Value * 7 + 5, 55, 400);
                SheetGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = ColumnName(c),
                    Width = width,
                    Binding = new Binding($"[{c}]") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.Explicit },
                    ElementStyle = textStyle,
                    EditingElementStyle = new Style(typeof(TextBox))
                    {
                        Setters = { new Setter(TextBox.PaddingProperty, new Thickness(5, 3, 4, 2)) }
                    }
                });
            }
            _rows.Clear();
            for (var r = 0; r < _rowCount; r++) _rows.Add(new SheetRow(this, r));
        }
        else
        {
            foreach (var row in _rows) row.Refresh();
        }
        UpdateFormulaBar();
        SelectionText.Text = $"{sheet.Rows.Count:N0} 行数据  ·  {(_document.ScriptsEnabled ? "QuickJS 已启用" : "缓存预览")}  ·  本地运行";
    }

    private void UpdateChrome()
    {
        FileTitle.Text = _displayName + (_dirty ? "  •" : "");
        Title = $"{(_dirty ? "* " : "")}{_displayName} — QuickExcel";
        FileSubtitle.Text = _filePath == null ? "本地工作簿 · 保存为 .xlsq 保留脚本" : _filePath;
        UndoButton.IsEnabled = _undo.Count > 0 && !_busy;
        RedoButton.IsEnabled = _redo.Count > 0 && !_busy;
    }

    private void UpdateFormulaBar()
    {
        CellAddress.Text = ColumnName(_selectedColumn) + (_selectedRow + 1);
        FormulaBar.Text = _document.GetCellInput(_sheetIndex, _selectedRow, _selectedColumn);
    }

    private async Task<bool> WorkAsync(string message, Action operation)
    {
        if (_busy) return false;
        _busy = true;
        Toolbar.IsEnabled = Workspace.IsEnabled = false;
        StatusText.Text = message;
        var timer = Stopwatch.StartNew();
        try
        {
            await Task.Run(operation);
            StatusText.Text = $"{message.TrimEnd('…')}完成 · {timer.ElapsedMilliseconds:N0} ms";
            return true;
        }
        catch (Exception ex)
        {
            StatusText.Text = "操作未完成";
            AppendLog(ex.Message);
            MessageBox.Show(this, ex.Message, "QuickExcel", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        finally
        {
            _busy = false;
            Toolbar.IsEnabled = Workspace.IsEnabled = true;
            UpdateChrome();
        }
    }

    private void ShowDiagnostics()
    {
        ScriptState.Text = _document.ScriptsEnabled ? "● 脚本已启用" : "● 脚本未启用";
        if (_document.ScriptErrors.Count > 0)
            AppendLog(string.Join("\n", _document.ScriptErrors.Take(8)));
        RefreshSheet();
    }

    private void AppendLog(string message)
    {
        var text = LogText.Text + $"\n[{DateTime.Now:HH:mm:ss}] {message}";
        LogText.Text = text.Length > 12000 ? text[^12000..] : text;
        LogText.ScrollToEnd();
    }

    private async Task RecalculateAsync()
    {
        if (_busy) return;
        _calculationPending = false;
        await WorkAsync("重新计算…", _document.Calculate);
        ShowDiagnostics();
    }

    private EditBatch CreateBatch(IEnumerable<(int Row, int Column, string Value)> values)
    {
        var edits = values.Select(v => new CellEdit(v.Row, v.Column,
            _document.GetCellInput(_sheetIndex, v.Row, v.Column), v.Value,
            _document.GetCellValue(_sheetIndex, v.Row, v.Column))).Where(e => e.Before != e.After).ToList();
        return new EditBatch(_sheetIndex, edits);
    }

    private void ApplyBatch(EditBatch batch, bool undo = false, bool record = true)
    {
        // Preflight the entire rectangle before the first mutation so a rejected paste is atomic.
        var sheet = _document.Workbook[batch.Sheet];
        long newCells = 0;
        foreach (var group in batch.Cells.GroupBy(c => c.Row))
        {
            if (group.Key < 0 || group.Key >= WorkbookDocument.MaxRows) throw new ArgumentOutOfRangeException(nameof(batch));
            var count = group.Max(c => c.Column) + 1;
            if (count > 256 || group.Any(c => c.Column < 0)) throw new ArgumentException("编辑范围不能超过 256 列。");
            newCells += Math.Max(0, count - (sheet.Rows.ElementAtOrDefault(group.Key)?.Cells.Count ?? 0));
        }
        if (_document.Workbook.Sum(s => s.Rows.Sum(r => (long)r.Cells.Count)) + newCells > WorkbookDocument.MaxCells)
            throw new InvalidOperationException("编辑后的工作簿超过原型的单元格容量限制。");
        if (batch.Cells.Any(c => (undo ? c.Before : c.After).Length > 32767))
            throw new ArgumentException("单元格最多支持 32767 个字符，整批编辑尚未应用。");
        foreach (var edit in batch.Cells)
        {
            if (undo) edit.AfterCache = _document.GetCellValue(batch.Sheet, edit.Row, edit.Column);
            _document.SetCell(batch.Sheet, edit.Row, edit.Column, undo ? edit.Before : edit.After);
            var cache = undo ? edit.BeforeCache : edit.AfterCache;
            var cell = sheet.Rows[edit.Row].Cells[edit.Column];
            if (cell.HasFormula && cache != null) cell.SetCalculatedValue(cache);
        }
        if (record && batch.Cells.Count > 0)
        {
            _undo.Push(batch);
            _redo.Clear();
        }
        if (batch.Cells.Count > 0) { _dirty = true; _calculationPending = true; }
        UpdateChrome();
    }

    private async Task ApplyFormulaBarAsync()
    {
        if (_busy) return;
        var batch = CreateBatch([(_selectedRow, _selectedColumn, FormulaBar.Text)]);
        ApplyBatch(batch);
        if (batch.Cells.Count > 0) await RecalculateAsync();
    }

    private async void ApplyCell_Click(object sender, RoutedEventArgs e) => await ApplyFormulaBarAsync();
    private async void FormulaBar_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) { e.Handled = true; await ApplyFormulaBarAsync(); SheetGrid.Focus(); }
        else if (e.Key == Key.Escape) { e.Handled = true; UpdateFormulaBar(); SheetGrid.Focus(); }
    }

    private void Grid_LoadingRow(object sender, DataGridRowEventArgs e)
    {
        if (e.Row.Item is SheetRow row) e.Row.Header = (row.Index + 1).ToString(CultureInfo.InvariantCulture);
    }

    private void Grid_CurrentCellChanged(object? sender, EventArgs e)
    {
        if (SheetGrid.CurrentCell.Item is not SheetRow row || SheetGrid.CurrentCell.Column == null) return;
        _selectedRow = row.Index;
        _selectedColumn = SheetGrid.CurrentCell.Column.DisplayIndex;
        UpdateFormulaBar();
    }

    private void Grid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
    {
        if (e.Row.Item is SheetRow row && e.EditingElement is TextBox editor)
        {
            editor.Text = _document.GetCellInput(_sheetIndex, row.Index, e.Column.DisplayIndex);
            editor.SelectAll();
        }
    }

    private void Grid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (_busy || e.EditAction != DataGridEditAction.Commit || e.Row.Item is not SheetRow row || e.EditingElement is not TextBox editor) return;
        var batch = CreateBatch([(row.Index, e.Column.DisplayIndex, editor.Text)]);
        try { ApplyBatch(batch); }
        catch (Exception ex) { e.Cancel = true; AppendLog(ex.Message); return; }
        if (batch.Cells.Count > 0)
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(async () =>
            {
                if (_calculationPending && !_busy && !_suppressAutoCalculation) await RecalculateAsync();
            }));
    }

    private async void Grid_KeyDown(object sender, KeyEventArgs e)
    {
        if (_busy || e.OriginalSource is TextBox) return;
        var ctrl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
        if (ctrl && e.Key == Key.V)
        {
            e.Handled = true;
            if (!Clipboard.ContainsText()) return;
            var input = Clipboard.GetText().Replace("\r\n", "\n").TrimEnd('\n');
            var values = new List<(int Row, int Column, string Value)>();
            var lines = input.Split('\n');
            for (var r = 0; r < lines.Length; r++)
            {
                var fields = lines[r].Split('\t');
                for (var c = 0; c < fields.Length; c++) values.Add((_selectedRow + r, _selectedColumn + c, fields[c]));
            }
            if (values.Count > 20000 || values.Any(v => v.Row >= 100000 || v.Column >= 256))
            {
                MessageBox.Show(this, "原型单次粘贴上限为 20,000 个单元格，编辑范围为 100,000 行 × 256 列。", "QuickExcel");
                return;
            }
            ApplyBatch(CreateBatch(values));
            RefreshSheet(true);
            await RecalculateAsync();
        }
        else if (ctrl && e.Key == Key.C)
        {
            e.Handled = true;
            var cells = SheetGrid.SelectedCells.Where(c => c.Item is SheetRow).ToList();
            if (cells.Count == 0) return;
            var minRow = cells.Min(c => ((SheetRow)c.Item).Index);
            var maxRow = cells.Max(c => ((SheetRow)c.Item).Index);
            var minCol = cells.Min(c => c.Column.DisplayIndex);
            var maxCol = cells.Max(c => c.Column.DisplayIndex);
            if ((long)(maxRow - minRow + 1) * (maxCol - minCol + 1) > 100000)
            { StatusText.Text = "复制范围过大，请缩小选择范围。"; return; }
            var sb = new StringBuilder();
            for (var r = minRow; r <= maxRow; r++)
            {
                if (r > minRow) sb.AppendLine();
                for (var c = minCol; c <= maxCol; c++)
                {
                    if (c > minCol) sb.Append('\t');
                    sb.Append(_document.GetCellValue(_sheetIndex, r, c));
                }
            }
            Clipboard.SetText(sb.ToString());
            StatusText.Text = "已复制计算结果。";
        }
        else if (e.Key == Key.Delete)
        {
            e.Handled = true;
            ApplyBatch(CreateBatch(SheetGrid.SelectedCells.Where(c => c.Item is SheetRow)
                .Select(c => (((SheetRow)c.Item).Index, c.Column.DisplayIndex, ""))));
            await RecalculateAsync();
        }
    }

    private async void Undo_Click(object sender, RoutedEventArgs e) => await UndoAsync(false);
    private async void Redo_Click(object sender, RoutedEventArgs e) => await UndoAsync(true);
    private async Task UndoAsync(bool redo)
    {
        if (_busy) return;
        var source = redo ? _redo : _undo;
        if (source.Count == 0) return;
        var batch = source.Pop();
        ApplyBatch(batch, undo: !redo, record: false);
        (redo ? _undo : _redo).Push(batch);
        if (_sheetIndex != batch.Sheet) SheetPicker.SelectedIndex = batch.Sheet;
        await RecalculateAsync();
    }

    private void SheetPicker_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_refreshing || SheetPicker.SelectedIndex < 0 || _busy) return;
        CommitEdit();
        _sheetIndex = SheetPicker.SelectedIndex;
        _selectedColumn = _selectedRow = 0;
        _columnCount = 12;
        _rowCount = 100;
        RefreshSheet(true);
    }

    private void MoreRows_Click(object sender, RoutedEventArgs e) { _rowCount = Math.Min(100000, _rowCount + 100); RefreshSheet(true); }
    private void MoreColumns_Click(object sender, RoutedEventArgs e) { _columnCount = Math.Min(256, _columnCount + 8); RefreshSheet(true); }
    private void AddSheet_Click(object sender, RoutedEventArgs e)
    {
        CommitEdit();
        var number = _document.Workbook.Worksheets.Count + 1;
        while (_document.Workbook.Worksheets.Any(s => s.Name == $"工作表{number}")) number++;
        _document.Workbook.NewWorksheet($"工作表{number}");
        _dirty = true;
        _refreshing = true;
        SheetPicker.ItemsSource = _document.Workbook.Worksheets.Select(s => s.Name).ToList();
        _refreshing = false;
        SheetPicker.SelectedIndex = _document.Workbook.Worksheets.Count - 1;
        UpdateChrome();
    }

    private void ScriptEditor_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_refreshing || ScriptState == null) return;
        _document.ScriptSource = ScriptEditor.Text;
        _document.DisableScripts();
        _dirty = true;
        ScriptState.Text = "● 已修改，等待应用";
        UpdateChrome();
    }

    private async void Run_Click(object sender, RoutedEventArgs e) => await RunScriptsAsync();
    private async Task RunScriptsAsync()
    {
        if (_busy) return;
        _suppressAutoCalculation = true;
        try { CommitEdit(); }
        finally { _suppressAutoCalculation = false; }
        _calculationPending = false;
        _document.ScriptSource = ScriptEditor.Text;
        var ok = await WorkAsync("应用 JavaScript 并计算…", _document.EnableScriptsAndCalculate);
        if (ok) { _dirty = true; AppendLog("脚本已应用，工作簿已重新计算。"); }
        ShowDiagnostics();
        UpdateChrome();
    }

    private void Disable_Click(object sender, RoutedEventArgs e)
    {
        _document.DisableScripts();
        ScriptState.Text = "● 脚本未启用";
        AppendLog("脚本已停用；自定义公式保留上次计算结果。");
        RefreshSheet();
    }

    private void CommitEdit()
    {
        SheetGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        SheetGrid.CommitEdit(DataGridEditingUnit.Row, true);
    }

    private async Task FlushEditsAsync()
    {
        _suppressAutoCalculation = true;
        try { CommitEdit(); }
        finally { _suppressAutoCalculation = false; }
        if (_calculationPending) await RecalculateAsync();
    }

    private async Task<bool> ConfirmDiscardAsync()
    {
        await FlushEditsAsync();
        if (!_dirty) return true;
        var result = MessageBox.Show(this, "保存当前工作簿和脚本后再继续？", "QuickExcel",
            MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        return result == MessageBoxResult.No || result == MessageBoxResult.Yes && await SaveAsync(false);
    }

    private async void New_Click(object sender, RoutedEventArgs e)
    {
        if (!_busy && await ConfirmDiscardAsync()) InstallDocument(WorkbookDocument.CreateBlank(), null, "未命名.xlsq");
    }
    private async void Demo_Click(object sender, RoutedEventArgs e)
    {
        if (!_busy && await ConfirmDiscardAsync()) InstallDocument(WorkbookDocument.CreateDemo(), null, "示例工作簿.xlsq");
    }
    private async void Open_Click(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        var dialog = new OpenFileDialog { Filter = "工作簿 (*.xlsq;*.xlsx)|*.xlsq;*.xlsx|QuickExcel (*.xlsq)|*.xlsq|Excel (*.xlsx)|*.xlsx", Title = "打开工作簿" };
        if (dialog.ShowDialog(this) == true && await ConfirmDiscardAsync()) await OpenPathAsync(dialog.FileName);
    }

    private async Task OpenPathAsync(string path)
    {
        WorkbookDocument? next = null;
        if (await WorkAsync("打开工作簿…", () => next = WorkbookDocument.Open(path)))
        {
            InstallDocument(next!, path, Path.GetFileName(path));
        }
    }

    private async void Save_Click(object sender, RoutedEventArgs e) => await SaveAsync(false);
    private async void SaveAs_Click(object sender, RoutedEventArgs e) => await SaveAsync(true);
    private async Task<bool> SaveAsync(bool saveAs)
    {
        if (_busy) return false;
        await FlushEditsAsync();
        string? path = saveAs ? null : _filePath;
        if (path == null)
        {
            var dialog = new SaveFileDialog { Filter = "QuickExcel 工作簿 (*.xlsq)|*.xlsq", DefaultExt = ".xlsq", AddExtension = true, FileName = Path.GetFileNameWithoutExtension(_displayName) + ".xlsq", Title = "保存工作簿和 JavaScript" };
            if (dialog.ShowDialog(this) != true) return false;
            path = dialog.FileName;
        }
        _document.ScriptSource = ScriptEditor.Text;
        if (!await WorkAsync("保存 .xlsq…", () => _document.Save(path))) return false;
        _filePath = path;
        _displayName = Path.GetFileName(path);
        _dirty = false;
        UpdateChrome();
        return true;
    }

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        await FlushEditsAsync();
        var dialog = new SaveFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", DefaultExt = ".xlsx", AddExtension = true, FileName = Path.GetFileNameWithoutExtension(_displayName) + "-结果.xlsx", Title = "导出当前结果（所有公式转为值，不含脚本）" };
        if (dialog.ShowDialog(this) == true && await WorkAsync("导出当前结果…", () => _document.ExportXlsx(dialog.FileName)))
            AppendLog("已导出 .xlsx：所有公式已转为当前结果，原工作簿和脚本保持完整。");
    }

    private async void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (_busy || !Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) return;
        switch (e.Key)
        {
            case Key.S: e.Handled = true; await SaveAsync(Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)); break;
            case Key.O: e.Handled = true; Open_Click(sender, e); break;
            case Key.N: e.Handled = true; New_Click(sender, e); break;
            case Key.Enter: e.Handled = true; await RunScriptsAsync(); break;
            case Key.Z when !ScriptEditor.IsKeyboardFocusWithin && !FormulaBar.IsKeyboardFocusWithin && e.OriginalSource is not TextBox:
                e.Handled = true; await UndoAsync(false); break;
            case Key.Y when !ScriptEditor.IsKeyboardFocusWithin && !FormulaBar.IsKeyboardFocusWithin && e.OriginalSource is not TextBox:
                e.Handled = true; await UndoAsync(true); break;
        }
    }

    private async void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_closingApproved) { _document.Dispose(); return; }
        if (_busy) { e.Cancel = true; StatusText.Text = "请等待当前操作结束。"; return; }
        e.Cancel = true;
        if (await ConfirmDiscardAsync()) { _closingApproved = true; Close(); }
    }

    public static string ColumnName(int index)
    {
        var result = "";
        do { result = (char)('A' + index % 26) + result; index = index / 26 - 1; } while (index >= 0);
        return result;
    }

    private async Task RunSmokeTestAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        // Drive the same editing, undo and formula-bar paths used by the window.
        InstallDocument(WorkbookDocument.CreateBlank(), null, "界面验收.xlsq");
        ScriptEditor.Text = "quickExcel.functions.register(\"DOUBLE\", x => x * 2);";
        ApplyBatch(CreateBatch([(0, 0, "21"), (0, 1, "=DOUBLE(A1)")]));
        await Task.Run(_document.EnableScriptsAndCalculate);
        if (_document.GetCellValue(0, 0, 1) != "42") throw new Exception("JS formula integration failed.");
        _selectedRow = 0; _selectedColumn = 0;
        FormulaBar.Text = "25";
        await ApplyFormulaBarAsync();
        if (_document.GetCellValue(0, 0, 1) != "50") throw new Exception("Formula-bar recalculation failed.");
        await UndoAsync(false);
        if (_document.GetCellValue(0, 0, 1) != "42") throw new Exception("Undo recalculation failed.");
        await UndoAsync(true);
        if (_document.GetCellValue(0, 0, 1) != "50") throw new Exception("Redo recalculation failed.");
        _document.DisableScripts();
        ApplyBatch(CreateBatch([(0, 1, "99")]));
        await UndoAsync(false);
        if (_document.GetCellValue(0, 0, 1) != "50") throw new Exception("Disabled formula cache was lost on undo.");
        var beforeRejected = _document.GetCellValue(0, 0, 0);
        try
        {
            ApplyBatch(CreateBatch([(0, 0, "999"), (1, 0, new string('x', 32768))]));
            throw new Exception("Oversized paste was accepted.");
        }
        catch (ArgumentException) { }
        if (_document.GetCellValue(0, 0, 0) != beforeRejected) throw new Exception("Rejected paste partially changed workbook.");
        await Task.Run(_document.EnableScriptsAndCalculate);
        RefreshSheet();
        var path = Path.Combine(folder, "smoke.xlsq");
        _filePath = path;
        SheetGrid.CurrentCell = new DataGridCellInfo(_rows[0], SheetGrid.Columns[0]);
        SheetGrid.ScrollIntoView(_rows[0], SheetGrid.Columns[0]);
        SheetGrid.Focus();
        if (!SheetGrid.BeginEdit()) throw new Exception("DataGrid could not enter edit mode.");
        if (SheetGrid.Columns[0].GetCellContent(_rows[0]) is not TextBox editBox) throw new Exception("DataGrid editor missing.");
        editBox.Text = "30";
        if (!await SaveAsync(false)) throw new Exception("Saving an active cell edit failed.");
        using (var reopened = WorkbookDocument.Open(path))
        {
            if (reopened.ScriptsEnabled) throw new Exception("Opened scripts ran automatically.");
            if (reopened.GetCellValue(0, 0, 1) != "60") throw new Exception("Saved formula cache did not include the pending edit.");
            reopened.EnableScriptsAndCalculate();
            if (reopened.GetCellValue(0, 0, 1) != "60") throw new Exception("Package roundtrip failed.");
            reopened.ExportXlsx(Path.Combine(folder, "smoke.xlsx"));
        }
        InstallDocument(WorkbookDocument.CreateDemo(), null, "示例工作簿.xlsq");
        await Task.Run(_document.EnableScriptsAndCalculate);
        ShowDiagnostics();
        AppendLog("界面验收通过：JS 计算、公式栏编辑、撤销/重做、.xlsq 往返、.xlsx 导出。");
        _document.Save(Path.Combine(folder, "QuickExcel-示例.xlsq"));
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        UpdateLayout();
        var dpi = VisualTreeHelper.GetDpi(this);
        var bitmap = new RenderTargetBitmap((int)(ActualWidth * dpi.DpiScaleX), (int)(ActualHeight * dpi.DpiScaleY), 96 * dpi.DpiScaleX, 96 * dpi.DpiScaleY, PixelFormats.Pbgra32);
        bitmap.Render(this);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using (var file = File.Create(Path.Combine(folder, "QuickExcel-preview.png"))) encoder.Save(file);
        File.WriteAllText(Path.Combine(folder, "smoke-result.txt"), "PASS: native QuickJS, formula-bar edit/recalculate, undo/redo with cached formulas, atomic rejected paste, save active DataGrid edit with fresh formula caches, XLSQ roundtrip, values-only XLSX export, WPF render.");
    }

    private sealed record CellEdit(int Row, int Column, string Before, string After, string BeforeCache)
    {
        public string? AfterCache { get; set; }
    }
    private sealed record EditBatch(int Sheet, List<CellEdit> Cells);

    private sealed class SheetRow(MainWindow owner, int index) : INotifyPropertyChanged
    {
        public int Index { get; } = index;
        public string this[int column] { get => owner._document.GetCellValue(owner._sheetIndex, Index, column); set { } }
        public XlStyle? GetStyle(int column)
        {
            var sheet = owner._document.Workbook.Worksheets[owner._sheetIndex];
            var row = sheet.Rows.ElementAtOrDefault(Index);
            return row?.Cells.ElementAtOrDefault(column)?.Style ?? row?.Style ?? (sheet.Columns.TryGetValue(column, out var col) ? col.Style : null);
        }
        public event PropertyChangedEventHandler? PropertyChanged;
        public void Refresh() => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Item[]"));
    }

    private sealed class AppearanceConverter(int column, string property) : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var style = (value as SheetRow)?.GetStyle(column);
            return property switch
            {
                "weight" => style?.Bold == true ? FontWeights.SemiBold : FontWeights.Normal,
                "background" => ColorBrush(style?.FillColor, Brushes.Transparent),
                _ => ColorBrush(style?.FontColor, new SolidColorBrush(Color.FromRgb(34, 53, 46)))
            };
        }
        private static Brush ColorBrush(string? color, Brush fallback)
        {
            try { return string.IsNullOrEmpty(color) ? fallback : (Brush)new BrushConverter().ConvertFromString("#" + color.TrimStart('#'))!; }
            catch { return fallback; }
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
    }
}
