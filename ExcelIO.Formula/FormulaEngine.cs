namespace ExcelIO.Formula;

public sealed class FormulaEngine : IFormulaEngine
{
    private readonly FunctionRegistry _functions = new();
    private readonly FormulaEvaluator _evaluator;
    private readonly List<CircularReference> _circularRefs = [];

    public IFunctionRegistry Functions => _functions;
    public IReadOnlyList<CircularReference> CircularReferences => _circularRefs;

    public FormulaEngine()
    {
        BuiltinFunctions.RegisterAll(_functions);
        _evaluator = new FormulaEvaluator(_functions);
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

        var context = new SheetFormulaContext(worksheet, worksheet.Workbook);

        for (int r = 0; r < worksheet.Rows.Count; r++)
        {
            var row = worksheet.Rows[r];
            for (int c = 0; c < row.Cells.Count; c++)
            {
                var cell = row.Cells[c];
                if (!cell.HasFormula) continue;

                var result = _evaluator.EvaluateFormula(cell.Formula!, context);
                cell.Value = CoerceResult(result);
            }
        }
    }

    public void Calculate(XlWorkbook workbook)
    {
        _circularRefs.Clear();
        foreach (var sheet in workbook.Worksheets)
            Calculate(sheet);
    }

    private static string CoerceResult(object result)
    {
        if (result is null) return "";
        if (result is double d)
        {
            // Avoid "15.0" for whole numbers
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

    /// <summary>
    /// Default IFormulaContext wrapping an XlWorksheet.
    /// </summary>
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
