using ExcelIO.Formula;

namespace ExcelIO.Formula.Test;

public class CircularReferenceTests
{
    [Fact]
    public void SelfReference_CellReferencesItself()
    {
        // A1 = A1+1
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var a1 = row.Insert(0, "");
        a1.SetFormula("=A1+1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.NotEmpty(engine.CircularReferences);
    }

    [Fact]
    public void DirectMutual_TwoCellsReferenceEachOther()
    {
        // A1 = B1+1, B1 = A1+1
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
        // Either A1 or B1 should be in a cycle
        var allCycleCells = engine.CircularReferences.SelectMany(c => c.Path).ToHashSet();
        Assert.Contains((0, 0), allCycleCells);
        Assert.Contains((0, 1), allCycleCells);
    }

    [Fact]
    public void ThreeNodeCycle_DaisyChain()
    {
        // A1 = B1+1, B1 = C1+2, C1 = A1+3
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var a1 = row.Insert(0, "");
        a1.SetFormula("=B1+1");
        var b1 = row.Insert(1, "");
        b1.SetFormula("=C1+2");
        var c1 = row.Insert(2, "");
        c1.SetFormula("=A1+3");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.NotEmpty(engine.CircularReferences);
        var allCycleCells = engine.CircularReferences.SelectMany(c => c.Path).ToHashSet();
        Assert.Contains((0, 0), allCycleCells);
        Assert.Contains((0, 1), allCycleCells);
        Assert.Contains((0, 2), allCycleCells);
    }

    [Fact]
    public void CyclePlusIndependent_IndependentCalculatesAnyway()
    {
        // A1=5
        // B1=A1  (independent → 5)
        // C1=D1+1, D1=C1+1  (cycle)
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("5");                       // A1
        var b1 = row.Insert(1, "");
        b1.SetFormula("=A1");               // B1 = A1 (independent)
        var c1 = row.Insert(2, "");
        c1.SetFormula("=D1+1");             // C1 = D1+1 (cycle)
        var d1 = row.Insert(3, "");
        d1.SetFormula("=C1+1");             // D1 = C1+1 (cycle)

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        // B1 is independent and should be calculated
        Assert.Equal("5", b1.Value);

        // Cycle should be reported
        Assert.NotEmpty(engine.CircularReferences);
        var allCycleCells = engine.CircularReferences.SelectMany(c => c.Path).ToHashSet();
        Assert.Contains((0, 2), allCycleCells); // C1
        Assert.Contains((0, 3), allCycleCells); // D1
        Assert.DoesNotContain((0, 1), allCycleCells); // B1 NOT in cycle
    }

    [Fact]
    public void TwoDisconnectedCycles_BothReported()
    {
        // A1=B1+1, B1=A1+1   (cycle 1)
        // C1=D1+1, D1=C1+1   (cycle 2)
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var a1 = row.Insert(0, "");
        a1.SetFormula("=B1+1");
        var b1 = row.Insert(1, "");
        b1.SetFormula("=A1+1");
        var c1 = row.Insert(2, "");
        c1.SetFormula("=D1+1");
        var d1 = row.Insert(3, "");
        d1.SetFormula("=C1+1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.NotEmpty(engine.CircularReferences);
        // Should have 2 cycle groups (might be merged if FindCycle finds a single path)
        // At minimum, all 4 cells should appear in reported cycles
        var allCycleCells = engine.CircularReferences.SelectMany(c => c.Path).ToHashSet();
        Assert.Contains((0, 0), allCycleCells);
        Assert.Contains((0, 1), allCycleCells);
        Assert.Contains((0, 2), allCycleCells);
        Assert.Contains((0, 3), allCycleCells);
    }

    [Fact]
    public void NoCycle_NormalFormulas_EmptyCircularReferences()
    {
        // A1=5, B1=A1*2, C1=B1+1  — linear chain, no cycle
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("5");
        var b1 = row.Insert(1, "");
        b1.SetFormula("=A1*2");
        var c1 = row.Insert(2, "");
        c1.SetFormula("=B1+1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.Empty(engine.CircularReferences);
        Assert.Equal("10", b1.Value);
        Assert.Equal("11", c1.Value);
    }

    [Fact]
    public void CycleIncludesCorrectCells()
    {
        // A1 = B1+1
        // B1 = C1+1
        // C1 = A1+1
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var a1 = row.Insert(0, "");
        a1.SetFormula("=B1+1");
        var b1 = row.Insert(1, "");
        b1.SetFormula("=C1+1");
        var c1 = row.Insert(2, "");
        c1.SetFormula("=A1+1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.Single(engine.CircularReferences);
        var cycle = engine.CircularReferences[0];
        // Cycle should contain at least 3 distinct cells + return to start
        var distinct = cycle.Path.Distinct().ToList();
        Assert.Equal(3, distinct.Count);
        Assert.Contains((0, 0), distinct);
        Assert.Contains((0, 1), distinct);
        Assert.Contains((0, 2), distinct);
    }

    [Fact]
    public void LargeDiamond_NoCycle_CorrectOrder()
    {
        // A1=10
        // B1=A1*2   (20)
        // C1=A1+5   (15)
        // D1=B1+C1  (35)
        // All acyclic — should calculate correctly
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        var b1 = row.Insert(1, "");
        b1.SetFormula("=A1*2");
        var c1 = row.Insert(2, "");
        c1.SetFormula("=A1+5");
        var d1 = row.Insert(3, "");
        d1.SetFormula("=B1+C1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.Empty(engine.CircularReferences);
        Assert.Equal("20", b1.Value);
        Assert.Equal("15", c1.Value);
        Assert.Equal("35", d1.Value);
    }

    [Fact]
    public void CycleViaRangeFormula_Detected()
    {
        // A1:A3 = 1, 2, 3
        // B1 = SUM(A1:A3) + C1
        // C1 = B1 * 2
        // B1 depends on C1, C1 depends on B1 → cycle
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        ws.NewRow(["1"]);
        ws.NewRow(["2"]);
        ws.NewRow(["3"]);
        var b1 = ws.Rows[0].Insert(1, "");
        b1.SetFormula("=SUM(A1:A3)+C1");
        var c1 = ws.Rows[0].Insert(2, "");
        c1.SetFormula("=B1*2");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.NotEmpty(engine.CircularReferences);
        var allCycleCells = engine.CircularReferences.SelectMany(c => c.Path).ToHashSet();
        Assert.Contains((0, 1), allCycleCells); // B1
        Assert.Contains((0, 2), allCycleCells); // C1
        // A1:A3 are input values, should NOT be in cycle
        Assert.DoesNotContain((0, 0), allCycleCells);
    }

    [Fact]
    public void FormulaReferencesNonExistentCell_HandlesGracefully()
    {
        // A1 = Z999+1  — Z999 is out of bounds
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var a1 = row.Insert(0, "");
        a1.SetFormula("=Z999+1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.Empty(engine.CircularReferences);
        // Should evaluate Z999 as 0, so result is 1
        Assert.Equal("1", a1.Value);
    }

    [Fact]
    public void SelfRef_DoesNotCrash()
    {
        // A1 = A1  (trivial self-ref, should not crash)
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var a1 = row.Insert(0, "");
        a1.SetFormula("=A1");

        var engine = new FormulaEngine();
        engine.Calculate(ws);

        Assert.NotEmpty(engine.CircularReferences);
    }
}
