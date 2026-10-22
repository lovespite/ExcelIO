using System.Globalization;
using System.Diagnostics;
using ExcelIO;
using ExcelIO.Formula;
using QuickExcel.Scripting;

namespace QuickExcel.Core;

/// <summary>A single-threaded editable workbook. Opening a document never runs its JavaScript.</summary>
public sealed class WorkbookDocument : IDisposable
{
    public const int MaxRows = 100_000;
    public const int MaxColumns = 1_024;
    public const int MaxCells = 1_000_000;
    public const string DefaultScript = """
        // 单元格：=DISCOUNT(B2,D2)*C2
        quickExcel.functions.register(
            "DISCOUNT", (price, rate) => {
                const result = price * (1 - rate);
                return Math.round(result * 100) / 100;
            }
        );

        // 范围传入数组：=TOTAL(B2:B4)
        quickExcel.functions.register(
            "TOTAL", values => values.reduce(
                (sum, value) => sum + Number(value),
                0
            )
        );
        """;

    private readonly List<string> _errors = [];
    private FunctionRegistry _functions = CreateBuiltinRegistry();
    private QuickJsSession? _session;
    private string _scriptSource;
    private bool _disposed;

    private WorkbookDocument(XlWorkbook workbook, string scriptSource)
    {
        Workbook = workbook;
        _scriptSource = scriptSource;
        if (Workbook.Count == 0) Workbook.NewWorksheet("Sheet1");
    }

    public XlWorkbook Workbook { get; }
    public bool ScriptsEnabled => _session is not null;
    public IReadOnlyList<string> ScriptErrors => _errors.AsReadOnly();

    /// <summary>Changing source invalidates the running script until explicitly enabled again.</summary>
    public string ScriptSource
    {
        get => _scriptSource;
        set
        {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(value);
            if (_scriptSource == value) return;
            DisableScripts();
            _scriptSource = value;
        }
    }

    public static WorkbookDocument CreateBlank() => new(new XlWorkbook(), DefaultScript);

    public static WorkbookDocument CreateDemo()
    {
        var document = CreateBlank();
        var sheet = document.Workbook[0];
        sheet.Name = "报价单";
        sheet.AddRow("商品", "单价", "数量", "折扣率", "实付金额");
        sheet.AddRow("机械键盘", "299", "2", "0.1", "538.2");
        sheet.AddRow("无线鼠标", "129", "3", "0.05", "367.65");
        sheet.AddRow("显示器支架", "189", "1", "0.15", "160.65");
        sheet.AddRow("合计", "", "", "", "1066.5");
        sheet.AddRow();
        sheet.AddRow("试用提示", "先启用 JavaScript，再修改单价或折扣率，查看金额重算。");
        sheet.AddRow("公式示例", "'=DISCOUNT(B2,D2)*C2");
        sheet.Rows[7].Cells[1].Value = "=DISCOUNT(B2,D2)*C2";
        for (var row = 1; row <= 3; row++)
            sheet.Rows[row].Cells[4].SetFormula($"=DISCOUNT(B{row + 1},D{row + 1})*C{row + 1}", sheet.Rows[row].Cells[4].Value);
        sheet.Rows[4].Cells[4].SetFormula("=SUM(E2:E4)", "1066.5");
        var header = new XlStyle { Bold = true, FillColor = "FF134E4A", FontColor = "FFFFFFFF" };
        foreach (var cell in sheet.Rows[0].Cells) cell.Style = header;
        foreach (var cell in sheet.Rows[4].Cells) cell.Style = new XlStyle { Bold = true, FillColor = "FFCCFBF1" };
        sheet.Columns[0] = new XlColumn { Width = 22 };
        for (var c = 1; c < 5; c++) sheet.Columns[c] = new XlColumn { Width = 18 };
        sheet.Options.TabColor = "FF0F766E";
        return document;
    }

    public static WorkbookDocument Open(string path)
    {
        var (workbook, script) = DocumentPackage.Open(path);
        return new WorkbookDocument(workbook, script);
    }

    public void Save(string path)
    {
        ThrowIfDisposed();
        DocumentPackage.Save(path, Workbook, ScriptSource);
    }

    /// <summary>Exports current displayed values; all formulas are replaced in a separate snapshot.</summary>
    public void ExportXlsx(string path)
    {
        ThrowIfDisposed();
        DocumentPackage.Export(path, Workbook);
    }

    public XlWorksheet AddSheet(string? name = null)
    {
        ThrowIfDisposed();
        if (name is null)
        {
            var number = Workbook.Count + 1;
            while (Workbook.WorksheetNames.Contains($"Sheet{number}", StringComparer.OrdinalIgnoreCase)) number++;
            name = $"Sheet{number}";
        }
        if (string.IsNullOrWhiteSpace(name) || name.Length > 31 || name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
            throw new ArgumentException("工作表名称应为 1–31 个字符，不能包含 []:*?/\\。", nameof(name));
        return Workbook.NewWorksheet(name);
    }

    public void EnableScriptsAndCalculate()
    {
        ThrowIfDisposed();
        DisableScripts();
        _errors.Clear();
        try
        {
            var session = new QuickJsSession(ScriptSource);
            _session = session;
            foreach (var name in session.FunctionNames)
            {
                _functions.Register(new ExcelFunction(name, "JavaScript", "工作簿中的 JavaScript 自定义函数", 0, int.MaxValue,
                    (arguments, _) => InvokeScript(session, name, arguments)));
            }
        }
        catch (Exception exception)
        {
            DisableScripts();
            AddError($"脚本未启用：{exception.Message}");
            throw;
        }
        Calculate();
    }

    public void DisableScripts()
    {
        ThrowIfDisposed();
        _session?.Dispose();
        _session = null;
        _functions = CreateBuiltinRegistry();
    }

    /// <summary>Recalculates explicit same-sheet dependencies. Unsupported cross-sheet formulas keep their cache.</summary>
    public void Calculate()
    {
        ThrowIfDisposed();
        _errors.Clear();
        // Excel's formula syntax and the stored numeric values use '.' independently of Windows locale.
        var previousCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            var timer = Stopwatch.StartNew();
            foreach (var sheet in Workbook)
            {
                if (timer.Elapsed > TimeSpan.FromSeconds(5))
                {
                    AddError("重算达到 5 秒限制；尚未计算的单元格保留上次结果。");
                    break;
                }
                CalculateSheet(sheet, timer);
            }
        }
        finally
        {
            CultureInfo.CurrentCulture = previousCulture;
        }
    }

    public void SetCell(int sheetIndex, int row, int column, string text)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(text);
        ValidateCoordinates(row, column);
        if (text.Length > 32_767) throw new ArgumentException("单元格最多支持 32767 个字符。", nameof(text));
        var sheet = Workbook[sheetIndex];
        var currentCount = Workbook.Sum(s => s.Rows.Sum(r => (long)r.Cells.Count));
        var existing = row < sheet.Rows.Count ? sheet.Rows[row].Cells.Count : 0;
        if (currentCount + Math.Max(0, column + 1 - existing) > MaxCells)
            throw new InvalidOperationException($"原型最多支持 {MaxCells:N0} 个单元格。");
        while (sheet.Rows.Count <= row) sheet.NewRow();
        var targetRow = sheet.Rows[row];
        while (targetRow.Cells.Count <= column) targetRow.Cells.Add(new XlCell(targetRow));
        var cell = targetRow.Cells[column];
        if (text.StartsWith('=')) cell.SetFormula(text);
        else cell.Value = text.StartsWith('\'') ? text[1..] : text;
    }

    public string GetCellInput(int sheetIndex, int row, int column)
    {
        var cell = FindCell(sheetIndex, row, column);
        if (cell is null) return "";
        if (cell.HasFormula) return cell.Formula!.StartsWith('=') ? cell.Formula : "=" + cell.Formula;
        // Preserve a literal that begins with '=' or an apostrophe when edited again.
        return cell.Value.StartsWith('=') || cell.Value.StartsWith('\'') ? "'" + cell.Value : cell.Value;
    }

    public string GetCellValue(int sheetIndex, int row, int column) => FindCell(sheetIndex, row, column)?.Value ?? "";

    private XlCell? FindCell(int sheetIndex, int row, int column)
    {
        ThrowIfDisposed();
        ValidateCoordinates(row, column);
        var sheet = Workbook[sheetIndex];
        return row < sheet.Count && column < sheet[row].Count ? sheet[row].Cells[column] : null;
    }

    private static void ValidateCoordinates(int row, int column)
    {
        if (row < 0 || row >= MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0 || column >= MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
    }

    private object InvokeScript(QuickJsSession session, string name, IReadOnlyList<object> arguments)
    {
        try
        {
            return session.Invoke(name, arguments) ?? "";
        }
        catch (Exception exception)
        {
            AddError($"{name}：{exception.Message}");
            return "#SCRIPT!";
        }
    }

    private void CalculateSheet(XlWorksheet sheet, Stopwatch timer)
    {
        var nodes = new Dictionary<(int Row, int Column), CalculationNode>();
        for (var row = 0; row < sheet.Count; row++)
        for (var column = 0; column < sheet[row].Count; column++)
        {
            if (timer.Elapsed > TimeSpan.FromSeconds(5))
            {
                AddError("重算达到 5 秒限制；尚未计算的单元格保留上次结果。");
                return;
            }
            var cell = sheet[row].Cells[column];
            if (!cell.HasFormula) continue;
            var node = new CalculationNode(cell);
            nodes[(row, column)] = node;
            if (HasCrossSheetReference(cell.Formula!))
            {
                node.PreserveCache = true;
                AddError("原型暂不重算跨工作表引用；这些单元格保留已保存的结果。");
                continue;
            }
            try
            {
                // Bound recursive parser depth before accepting externally authored formulas.
                if (cell.Formula!.Length > 8_192 || cell.Formula.Count(c => c is '(' or '+' or '-' or '^') > 128)
                    throw new FormatException("公式过长或嵌套过深。");
                node.Expression = FormulaParser.Parse(cell.Formula);
                Visit(node.Expression, expression =>
                {
                    if (expression is FunctionExpr function && !_functions.Exists(function.Name))
                    {
                        if (!ScriptsEnabled) node.PreserveCache = true;
                        AddError(ScriptsEnabled
                            ? $"未找到函数 {function.Name}。"
                            : $"函数 {function.Name} 尚未启用；相关单元格保留缓存结果，可能已过期。");
                    }
                });
            }
            catch (Exception exception)
            {
                node.Expression = new ErrorExpr("#VALUE!");
                AddError($"{sheet.Name}!{Address(row, column)}：{exception.Message}");
            }
        }

        // Only actual formula cells need graph edges. Large ranges never expand to empty coordinates.
        var byRow = nodes.Keys.GroupBy(p => p.Row).ToDictionary(g => g.Key, g => g.ToArray());
        foreach (var (position, node) in nodes)
        {
            if (timer.Elapsed > TimeSpan.FromSeconds(5))
            {
                AddError("重算达到 5 秒限制；尚未计算的单元格保留上次结果。");
                return;
            }
            var references = new HashSet<(int Row, int Column)>();
            if (node.Expression is not null) Visit(node.Expression, expression =>
            {
                if (expression is CellRefExpr reference && nodes.ContainsKey((reference.Row, reference.Col)))
                    references.Add((reference.Row, reference.Col));
                if (expression is RangeRefExpr range)
                {
                    var first = Math.Max(0, range.StartRow);
                    var last = Math.Min(sheet.Count - 1, range.EndRow);
                    for (var row = first; row <= last; row++)
                        if (byRow.TryGetValue(row, out var positions))
                            foreach (var candidate in positions)
                                if (candidate.Column >= range.StartCol && candidate.Column <= range.EndCol) references.Add(candidate);
                }
            });
            node.Pending = references.Count;
            foreach (var reference in references) nodes[reference].Dependents.Add(position);
        }

        var ready = new Queue<(int Row, int Column)>(nodes.Where(pair => pair.Value.Pending == 0).Select(pair => pair.Key));
        var evaluator = new FormulaEvaluator(_functions);
        var context = new SheetContext(sheet);
        while (ready.TryDequeue(out var position))
        {
            if (timer.Elapsed > TimeSpan.FromSeconds(5))
            {
                AddError("重算达到 5 秒限制；尚未计算的单元格保留上次结果。");
                return;
            }
            var node = nodes[position];
            node.Processed = true;
            if (!node.PreserveCache)
            {
                try { node.Cell.SetCalculatedValue(ToCellText(evaluator.Evaluate(node.Expression!, context))); }
                catch (Exception exception)
                {
                    node.Cell.SetCalculatedValue("#VALUE!");
                    AddError($"{sheet.Name}!{Address(position.Row, position.Column)}：{exception.Message}");
                }
            }
            foreach (var dependent in node.Dependents)
            {
                var next = nodes[dependent];
                next.PreserveCache |= node.PreserveCache;
                if (--next.Pending == 0) ready.Enqueue(dependent);
            }
        }
        foreach (var (position, node) in nodes.Where(pair => !pair.Value.Processed))
        {
            node.Cell.SetCalculatedValue("#CYCLE!");
            AddError($"{sheet.Name}!{Address(position.Row, position.Column)}：存在循环引用或依赖于循环引用。");
        }
    }

    private static string ToCellText(object? value) => value switch
    {
        null => "",
        bool boolean => boolean ? "TRUE" : "FALSE",
        double number when !double.IsFinite(number) => "#NUM!",
        double number => number.ToString("G15", CultureInfo.InvariantCulture),
        IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
        _ => value.ToString() ?? ""
    };

    private static void Visit(Expr expression, Action<Expr> action)
    {
        var pending = new Stack<Expr>();
        pending.Push(expression);
        while (pending.TryPop(out var current))
        {
            if (current is null) continue;
            action(current);
            switch (current)
            {
                case BinaryExpr binary: pending.Push(binary.Left); pending.Push(binary.Right); break;
                case UnaryExpr unary: pending.Push(unary.Operand); break;
                case FunctionExpr function: foreach (var argument in function.Arguments) pending.Push(argument); break;
            }
        }
    }

    private static bool HasCrossSheetReference(string formula)
    {
        var quoted = false;
        foreach (var character in formula)
        {
            if (character == '"') quoted = !quoted;
            if (!quoted && character == '!') return true;
        }
        return false;
    }

    private static FunctionRegistry CreateBuiltinRegistry()
    {
        var engine = new FormulaEngine();
        // The existing engine subscribes a no-op instance handler to a public static hook.
        // We only use its registered builtins; detach precisely its own handler immediately.
        if (XlCell.OnValueChanged is { } callbacks)
            foreach (var callback in callbacks.GetInvocationList())
                if (ReferenceEquals(callback.Target, engine)) XlCell.OnValueChanged -= (Action<XlCell>)callback;
        var registry = new FunctionRegistry();
        foreach (var function in engine.Functions.GetAll()) registry.Register(function);
        return registry;
    }

    private static string Address(int row, int column)
    {
        var letters = "";
        for (var value = column + 1; value > 0; value = (value - 1) / 26) letters = (char)('A' + (value - 1) % 26) + letters;
        return letters + (row + 1).ToString(CultureInfo.InvariantCulture);
    }

    private void AddError(string message)
    {
        if (_errors.Count < 100 && !_errors.Contains(message)) _errors.Add(message);
    }

    private void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);

    public void Dispose()
    {
        if (_disposed) return;
        _session?.Dispose();
        _session = null;
        _disposed = true;
    }

    private sealed class CalculationNode(XlCell cell)
    {
        public XlCell Cell { get; } = cell;
        public Expr? Expression { get; set; }
        public bool PreserveCache { get; set; }
        public bool Processed { get; set; }
        public int Pending { get; set; }
        public List<(int Row, int Column)> Dependents { get; } = [];
    }

    private sealed class SheetContext(XlWorksheet sheet) : IFormulaContext
    {
        public XlWorksheet Worksheet => sheet;
        public XlCell? GetCell(int row, int col) => row >= 0 && row < sheet.Count && col >= 0 && col < sheet[row].Count ? sheet[row].Cells[col] : null;
        public XlWorksheet? GetSheet(string name) => sheet.Workbook.Worksheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
    }
}
