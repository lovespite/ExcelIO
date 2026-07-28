using ExcelIO.Formula;

namespace ExcelIO.Formula.Test;

public class FormulaEvaluatorTests
{
    private readonly FormulaEngine _engine = new();

    [Fact]
    public void Evaluate_SimpleFormula_Calculates()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var cell = row.Insert(0, "");
        cell.SetFormula("=1+2+3");

        var result = _engine.Evaluate(cell, new TestContext(ws, wb));
        Assert.Equal("6", result);
    }

    [Fact]
    public void Evaluate_CellReference_GetsValue()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        row.Add("20");

        var cell = row.Insert(2, "");
        cell.SetFormula("=A1+B1");

        var result = _engine.Evaluate(cell, new TestContext(ws, wb));
        Assert.Equal("30", result);
    }

    [Fact]
    public void Evaluate_SumFunction_Calculates()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        row.Add("20");
        row.Add("30");

        var cell = row.Insert(3, "");
        cell.SetFormula("=SUM(A1:C1)");

        var result = _engine.Evaluate(cell, new TestContext(ws, wb));
        Assert.Equal("60", result);
    }

    [Fact]
    public void Evaluate_IfFunction_TrueBranch()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("15");

        var cell = row.Insert(1, "");
        cell.SetFormula("=IF(A1>10,\"big\",\"small\")");

        var result = _engine.Evaluate(cell, new TestContext(ws, wb));
        Assert.Equal("big", result);
    }

    [Fact]
    public void Evaluate_IfFunction_FalseBranch()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("5");

        var cell = row.Insert(1, "");
        cell.SetFormula("=IF(A1>10,\"big\",\"small\")");

        var result = _engine.Evaluate(cell, new TestContext(ws, wb));
        Assert.Equal("small", result);
    }

    [Fact]
    public void Evaluate_AverageFunction()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        row.Add("20");
        row.Add("30");

        var cell = row.Insert(3, "");
        cell.SetFormula("=AVERAGE(A1:C1)");

        var result = _engine.Evaluate(cell, new TestContext(ws, wb));
        Assert.Equal("20", result);
    }

    [Fact]
    public void Evaluate_MinMaxFunctions()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        row.Add("20");
        row.Add("5");

        var minCell = row.Insert(3, "");
        minCell.SetFormula("=MIN(A1:C1)");
        var maxCell = row.Insert(4, "");
        maxCell.SetFormula("=MAX(A1:C1)");

        Assert.Equal("5", _engine.Evaluate(minCell, new TestContext(ws, wb)));
        Assert.Equal("20", _engine.Evaluate(maxCell, new TestContext(ws, wb)));
    }

    [Fact]
    public void Evaluate_Concatenation()
    {
        var cell = new XlCell(new XlRow(new XlWorksheet(new XlWorkbook())));
        cell.SetFormula("=\"Hello\"&\" \"&\"World\"");
        var result = _engine.Evaluate(cell, new TestContext(null!, null!));
        Assert.Equal("Hello World", result);
    }

    [Fact]
    public void Evaluate_Comparison()
    {
        var cell = new XlCell(new XlRow(new XlWorksheet(new XlWorkbook())));
        cell.SetFormula("=10>5");
        Assert.Equal("TRUE", _engine.Evaluate(cell, new TestContext(null!, null!)));
        cell.SetFormula("=10<5");
        Assert.Equal("FALSE", _engine.Evaluate(cell, new TestContext(null!, null!)));
    }

    [Fact]
    public void Evaluate_DivideByZero_ReturnsError()
    {
        var cell = new XlCell(new XlRow(new XlWorksheet(new XlWorkbook())));
        cell.SetFormula("=1/0");
        var result = _engine.Evaluate(cell, new TestContext(null!, null!));
        Assert.StartsWith("#", result);
    }

    [Fact]
    public void Evaluate_EmptyCellRef_ReturnsZero()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");

        // B1 doesn't exist
        var cell = row.Insert(1, "");
        cell.SetFormula("=A1+B1");

        var result = _engine.Evaluate(cell, new TestContext(ws, wb));
        Assert.Equal("10", result); // B1 treated as 0
    }

    [Fact]
    public void Evaluate_TextFunctions()
    {
        var cell = new XlCell(new XlRow(new XlWorksheet(new XlWorkbook())));
        cell.SetFormula("=LEFT(\"hello\",2)");
        Assert.Equal("he", _engine.Evaluate(cell, new TestContext(null!, null!)));
        cell.SetFormula("=RIGHT(\"hello\",2)");
        Assert.Equal("lo", _engine.Evaluate(cell, new TestContext(null!, null!)));
        cell.SetFormula("=MID(\"hello\",2,3)");
        Assert.Equal("ell", _engine.Evaluate(cell, new TestContext(null!, null!)));
        cell.SetFormula("=LEN(\"hello\")");
        Assert.Equal("5", _engine.Evaluate(cell, new TestContext(null!, null!)));
        cell.SetFormula("=UPPER(\"hello\")");
        Assert.Equal("HELLO", _engine.Evaluate(cell, new TestContext(null!, null!)));
        cell.SetFormula("=TRIM(\" a b \")");
        Assert.Equal("a b", _engine.Evaluate(cell, new TestContext(null!, null!)));
    }

    [Fact]
    public void Evaluate_CountAndCountA()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        row.Add("hello");
        row.Add("20");
        row.Add("");

        // COUNT counts numbers only
        var c1 = row.Insert(4, "");
        c1.SetFormula("=COUNT(A1:D1)");
        Assert.Equal("2", _engine.Evaluate(c1, new TestContext(ws, wb)));

        // COUNTA counts non-empty
        var c2 = row.Insert(5, "");
        c2.SetFormula("=COUNTA(A1:D1)");
        Assert.Equal("3", _engine.Evaluate(c2, new TestContext(ws, wb)));
    }

    [Fact]
    public void Evaluate_Calculate_Worksheet()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");           // A1
        row.Add("20");           // B1
        var c = row.Insert(2, "");
        c.SetFormula("=A1+B1");  // C1 = A1+B1

        _engine.Calculate(ws);

        Assert.Equal("30", c.Value);
    }

    [Fact]
    public void Evaluate_VLookup()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var r1 = ws.NewRow(["apple", "1.5"]);
        var r2 = ws.NewRow(["banana", "2.0"]);
        var r3 = ws.NewRow(["cherry", "3.0"]);

        var cell = ws.NewRow().Insert(0, "");
        cell.SetFormula("=VLOOKUP(\"banana\",A1:B3,2)");

        var result = _engine.Evaluate(cell, new TestContext(ws, wb));
        Assert.Equal("2.0", result);
    }

    [Fact]
    public void Evaluate_CustomFunction_Works()
    {
        _engine.Functions.Register(new ExcelFunction(
            "DOUBLE", "Custom", "Double", 1, 1,
            (args, ctx) => Convert.ToDouble(args[0]) * 2));

        var cell = new XlCell(new XlRow(new XlWorksheet(new XlWorkbook())));
        cell.SetFormula("=DOUBLE(21)");

        var result = _engine.Evaluate(cell, new TestContext(null!, null!));
        Assert.Equal("42", result);
    }

    private sealed class TestContext : IFormulaContext
    {
        private readonly XlWorksheet _sheet;
        private readonly XlWorkbook _workbook;
        public XlWorksheet Worksheet => _sheet;

        public TestContext(XlWorksheet sheet, XlWorkbook workbook)
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
            => _workbook.Worksheets.FirstOrDefault(s =>
                string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
    }
}
