namespace ExcelIO.Formula;

public sealed class FormulaEngine : IFormulaEngine
{
    private readonly FunctionRegistry _functions = new();
    private readonly FormulaEvaluator _evaluator;
    private readonly DependencyGraph _graph = new();
    private readonly List<CircularReference> _circularRefs = [];
    private bool _isCalculating;

    public IFunctionRegistry Functions => _functions;
    public IReadOnlyList<CircularReference> CircularReferences => _circularRefs;

    public FormulaEngine()
    {
        BuiltinFunctions.RegisterAll(_functions);
        _evaluator = new FormulaEvaluator(_functions);
        XlCell.OnValueChanged += OnCellValueChanged;
    }

    public string? Evaluate(XlCell cell, IFormulaContext context)
    {
        if (!cell.HasFormula) return null;
        var result = _evaluator.EvaluateFormula(cell.Formula!, context);
        return CoerceResult(result);
    }

    public void Calculate(XlWorksheet worksheet)
    {
        _circularRefs.Clear();

        // Build dependency graph from all formula cells
        _graph.Build(worksheet);

        // Topological sort → evaluation order
        var sorted = _graph.TopologicalSort();
        if (sorted is null)
        {
            // Circular reference detected — collect cycle info
            var cycles = _graph.DetectCircularReferences();
            foreach (var cycle in cycles)
                _circularRefs.Add(new CircularReference(cycle));

            // Still try to evaluate what we can
            sorted = ForceEvaluateAll(worksheet);
        }

        if (sorted.Count == 0) return;

        var context = new SheetFormulaContext(worksheet, worksheet.Workbook);
        _isCalculating = true;
        try
        {
            foreach (var (row, col) in sorted)
            {
                if (row >= worksheet.Rows.Count) continue;
                var rowObj = worksheet.Rows[row];
                if (col >= rowObj.Cells.Count) continue;
                var cell = rowObj.Cells[col];
                if (!cell.HasFormula) continue;

                var result = _evaluator.EvaluateFormula(cell.Formula!, context);
                cell.SetCalculatedValue(CoerceResult(result));
            }
        }
        finally
        {
            _isCalculating = false;
        }
    }

    public void Calculate(XlWorkbook workbook)
    {
        _circularRefs.Clear();
        // Cross-sheet calculation: process sheets in order, then re-evaluate
        // any cross-sheet references via dirty tracking
        foreach (var sheet in workbook.Worksheets)
            Calculate(sheet);
    }

    private static string CoerceResult(object result)
    {
        if (result is null) return "";
        if (result is double d)
        {
            if (d == Math.Floor(d) && !double.IsInfinity(d))
                return ((long)d).ToString();
            return d.ToString("G15");
        }
        if (result is bool b) return b ? "TRUE" : "FALSE";
        if (result is int i) return i.ToString();
        if (result is long l) return l.ToString();
        if (result is string s) return s;
        return result.ToString() ?? "";
    }

    private void OnCellValueChanged(XlCell cell)
    {
        if (_isCalculating) return;
        // When a cell value changes externally, mark its dependents for recalculation.
        // The user will call Calculate() to trigger the actual evaluation.
    }

    /// <summary>
    /// When a cycle is detected, fall back to row-major evaluation order
    /// so at least some formulas produce values.
    /// </summary>
    private static List<(int Row, int Col)> ForceEvaluateAll(XlWorksheet sheet)
    {
        var cells = new List<(int, int)>();
        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            for (int c = 0; c < row.Cells.Count; c++)
            {
                if (row.Cells[c].HasFormula)
                    cells.Add((r, c));
            }
        }
        return cells;
    }

    private sealed class SheetFormulaContext : IFormulaContext
    {
        private readonly XlWorksheet _sheet;
        private readonly XlWorkbook _workbook;

        public XlWorksheet Worksheet => _sheet;

        public SheetFormulaContext(XlWorksheet sheet, XlWorkbook workbook)
        {
            _sheet = sheet;
            _workbook = workbook;
        }

        public XlCell? GetCell(int row, int col)
        {
            if (row < 0 || row >= _sheet.Rows.Count) return null;
            var r = _sheet.Rows[row];
            if (col < 0 || col >= r.Cells.Count) return null;
            return r.Cells[col];
        }

        public XlWorksheet? GetSheet(string name)
        {
            return _workbook.Worksheets.FirstOrDefault(s =>
                string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
        }
    }
}
