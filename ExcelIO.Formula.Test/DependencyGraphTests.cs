using ExcelIO.Formula;

namespace ExcelIO.Formula.Test;

public class DependencyGraphTests
{
    [Fact]
    public void SimpleLinearDependency_OrdersCorrectly()
    {
        // A1=10, B1=A1*2, C1=B1+5
        // Order: A1(no formula), B1, C1  → but only formula cells get sorted: B1, C1
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");                   // A1
        var b1 = row.Insert(1, "");
        b1.SetFormula("=A1*2");          // B1 depends on A1
        var c1 = row.Insert(2, "");
        c1.SetFormula("=B1+5");          // C1 depends on B1

        var graph = new DependencyGraph();
        graph.Build(ws);
        var sorted = graph.TopologicalSort();

        Assert.NotNull(sorted);
        // B1(row=0,col=1) must come before C1(row=0,col=2)
        var b1Idx = sorted!.IndexOf((0, 1));
        var c1Idx = sorted!.IndexOf((0, 2));
        Assert.True(b1Idx < c1Idx);
    }

    [Fact]
    public void DiamondDependency_AllPrecedentsBeforeDependent()
    {
        // A1=10, B1=20
        // C1=A1+B1  (C1 depends on A1, B1)
        // D1=C1*2  (D1 depends on C1)
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");                   // A1
        row.Add("20");                   // B1
        var c1 = row.Insert(2, "");
        c1.SetFormula("=A1+B1");         // C1
        var d1 = row.Insert(3, "");
        d1.SetFormula("=C1*2");          // D1

        var graph = new DependencyGraph();
        graph.Build(ws);
        var sorted = graph.TopologicalSort();

        Assert.NotNull(sorted);
        var c1Idx = sorted!.IndexOf((0, 2));
        var d1Idx = sorted!.IndexOf((0, 3));
        Assert.True(c1Idx < d1Idx);
    }

    [Fact]
    public void NoFormulas_ReturnsEmptyOrder()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        ws.NewRow(["a", "b", "c"]);

        var graph = new DependencyGraph();
        graph.Build(ws);
        var sorted = graph.TopologicalSort();

        Assert.NotNull(sorted);
        Assert.Empty(sorted!);
    }

    [Fact]
    public void IndependentFormulas_AnyOrderIsFine()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        var a2 = row.Insert(1, "");
        a2.SetFormula("=A1+1");
        var a3 = row.Insert(2, "");
        a3.SetFormula("=A1+2");  // independent of A2

        var graph = new DependencyGraph();
        graph.Build(ws);
        var sorted = graph.TopologicalSort();

        Assert.NotNull(sorted);
        Assert.Equal(2, sorted!.Count);
        Assert.Contains((0, 1), sorted);
        Assert.Contains((0, 2), sorted);
    }

    [Fact]
    public void RangeDependency_CreatesCorrectPrecedents()
    {
        // A1:A3 = 10, 20, 30
        // B1 = SUM(A1:A3)  → B1 depends on A1, A2, A3
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        ws.NewRow(["10", ""]);
        ws.NewRow(["20", ""]);
        ws.NewRow(["30", ""]);
        var b1 = ws.Rows[0].Cells[1];
        b1.SetFormula("=SUM(A1:A3)");

        var graph = new DependencyGraph();
        graph.Build(ws);

        // B1 should have dependents from A1, A2, A3
        var depsOfA1 = graph.GetDependents(0, 0); // A1
        Assert.Contains((0, 1), depsOfA1); // B1 depends on A1
    }

    [Fact]
    public void Calculate_LinearDependency_ProducesCorrectValues()
    {
        // A1=10, B1=A1*2, C1=B1+5  → C1 should be 25
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        var b1 = row.Insert(1, "");
        b1.SetFormula("=A1*2");
        var c1 = row.Insert(2, "");
        c1.SetFormula("=B1+5");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.Equal("20", b1.Value);
        Assert.Equal("25", c1.Value);
    }

    [Fact]
    public void Calculate_CascadingUpdate_AllLevelsEvaluated()
    {
        // A1=5
        // B1=A1+1     → 6
        // C1=B1*2     → 12
        // D1=C1+A1    → 17
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("5");
        var b1 = row.Insert(1, "");
        b1.SetFormula("=A1+1");
        var c1 = row.Insert(2, "");
        c1.SetFormula("=B1*2");
        var d1 = row.Insert(3, "");
        d1.SetFormula("=C1+A1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.Equal("6", b1.Value);
        Assert.Equal("12", c1.Value);
        Assert.Equal("17", d1.Value);
    }

    [Fact]
    public void Calculate_RangeFormula_EvaluatesCorrectly()
    {
        // A1:A3 = 10, 20, 30
        // B1 = SUM(A1:A3)
        // C1 = AVERAGE(A1:A3)
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        ws.NewRow(["10"]);
        ws.NewRow(["20"]);
        ws.NewRow(["30"]);
        var row0 = ws.Rows[0];
        var b1 = row0.Insert(1, "");
        b1.SetFormula("=SUM(A1:A3)");
        var c1 = row0.Insert(2, "");
        c1.SetFormula("=AVERAGE(A1:A3)");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.Equal("60", b1.Value);
        Assert.Equal("20", c1.Value);
    }

    [Fact]
    public void Calculate_MultiRowFormulas_OrdersCorrectly()
    {
        // Row1: A1=10, B1=A1+1
        // Row2: A2=B1+1, B2=A2*2
        // Row3: A3=B2+1
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var r1 = ws.NewRow(["10"]);
        var r1b = r1.Insert(1, "");
        r1b.SetFormula("=A1+1");
        var r2 = ws.NewRow([""]);
        var r2a = r2.Cells[0];
        r2a.SetFormula("=B1+1");
        var r2b = r2.Insert(1, "");
        r2b.SetFormula("=A2*2");
        var r3 = ws.NewRow([""]);
        var r3a = r3.Cells[0];
        r3a.SetFormula("=B2+1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.Equal("11", r1b.Value);     // 10+1
        Assert.Equal("12", r2a.Value);     // 11+1
        Assert.Equal("24", r2b.Value);     // 12*2
        Assert.Equal("25", r3a.Value);     // 24+1
    }

    [Fact]
    public void TopologicalSort_HasExpectedCount()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("1");
        row.Add("2");
        var c1 = row.Insert(2, "");
        c1.SetFormula("=A1+B1");
        var d1 = row.Insert(3, "");
        d1.SetFormula("=C1*3");

        var graph = new DependencyGraph();
        graph.Build(ws);
        var sorted = graph.TopologicalSort();

        Assert.NotNull(sorted);
        Assert.Equal(2, sorted!.Count); // 2 formula cells
    }

    [Fact]
    public void Calculate_CircularReference_ReportsCycles()
    {
        // A1 = B1+1, B1 = A1+1 → circular
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var a1 = row.Insert(0, "");
        a1.SetFormula("=B1+1");
        var b1 = row.Insert(1, "");
        b1.SetFormula("=A1+1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.NotEmpty(engine.CircularReferences);
    }

    [Fact]
    public void Calculate_BrokenDependency_AfterValueChange_RequiresRecalc()
    {
        // A1=10, B1=A1*2 → B1=20
        // Change A1=5, recalc → B1=10
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        var b1 = row.Insert(1, "");
        b1.SetFormula("=A1*2");

        var engine = new FormulaEngine();
        engine.Calculate(ws);
        Assert.Equal("20", b1.Value);

        // Change source value
        row.Cells[0].Value = "5";
        engine.Calculate(ws);

        Assert.Equal("10", b1.Value);
    }
}
