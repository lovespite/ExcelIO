namespace ExcelIO.Formula.Test;

public class SharedFormulaOptimizerTests
{
    [Fact]
    public void NoFormulas_ReturnsNull()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        ws.NewRow(["a", "b", "c"]);

        var map = XlSharedFormulaOptimizer.Build(ws);
        Assert.Null(map);
    }

    [Fact]
    public void SingleFormula_ReturnsNull()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        var c = row.Insert(1, "");
        c.SetFormula("=A1*2");

        var map = XlSharedFormulaOptimizer.Build(ws);
        Assert.Null(map); // need at least 2 shareable cells
    }

    [Fact]
    public void VerticalSharing_SameColumn()
    {
        // A1=10, A2=20, A3=30
        // B1=A1*2, B2=A2*2, B3=A3*2  → shareable vertically
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        ws.NewRow(["10"]);
        ws.NewRow(["20"]);
        ws.NewRow(["30"]);

        ws.Rows[0].Insert(1, "").SetFormula("=A1*2");
        ws.Rows[1].Insert(1, "").SetFormula("=A2*2");
        ws.Rows[2].Insert(1, "").SetFormula("=A3*2");

        var map = XlSharedFormulaOptimizer.Build(ws);
        Assert.NotNull(map);
        Assert.True(map!.Count >= 2);

        // B1 should be the master
        Assert.True(map.TryGetValue((0, 1), out var b1) && b1!.IsMaster);
        Assert.NotNull(b1.Ref);
        Assert.NotNull(b1.Formula);

        // B2, B3 should be children
        Assert.True(map.TryGetValue((1, 1), out var b2) && !b2!.IsMaster);
        Assert.True(map.TryGetValue((2, 1), out var b3) && !b3!.IsMaster);
    }

    [Fact]
    public void VerticalSharing_SkipsNonShareableRow()
    {
        // B1=A1*2, B2=DIFFERENT, B3=A3*2  → B1 and B3 are shareable but B2 breaks the run
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        ws.NewRow(["10"]);
        ws.NewRow(["20"]);
        ws.NewRow(["30"]);

        ws.Rows[0].Insert(1, "").SetFormula("=A1*2");
        ws.Rows[1].Insert(1, "").SetFormula("=A2+100"); // different formula
        ws.Rows[2].Insert(1, "").SetFormula("=A3*2");

        var map = XlSharedFormulaOptimizer.Build(ws);
        // B1 and B3 might group if consecutive shareable rows >= 2, but B2 breaks the pattern
        // Vertical scan: B1-B2-B3 run is 3 cells. Check B1 vs B2 → diff pattern → B2 excluded.
        // B1 alone can't be a group. B3 alone can't either. So no group.
        Assert.Null(map);
    }

    [Fact]
    public void HorizontalSharing_SameRow()
    {
        // A1=10, B1=A1+1, C1=B1+1  → shareable horizontally (B1 and C1 both have =X+1 pattern)
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        row.Add("10");
        row.Insert(1, "").SetFormula("=A1+1");
        row.Insert(2, "").SetFormula("=B1+1");

        var map = XlSharedFormulaOptimizer.Build(ws);
        // Horizontal: B1(0,1) =A1+1, C1(0,2) =B1+1
        // TranslateSharedFormula("=A1+1", 0, 1) → "=B1+1" ✓
        Assert.NotNull(map);
        Assert.True(map!.Count >= 2);
        Assert.True(map.TryGetValue((0, 1), out var b1) && b1!.IsMaster);
        Assert.True(map.TryGetValue((0, 2), out var c1) && !c1!.IsMaster);
    }

    [Fact]
    public void AbsoluteReferences_Shareable()
    {
        // B1=$A$1*2, B2=$A$1*2 → shareable (absolute ref doesn't change)
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        ws.NewRow(["10"]);
        ws.NewRow(["10"]);

        ws.Rows[0].Insert(1, "").SetFormula("=$A$1*2");
        ws.Rows[1].Insert(1, "").SetFormula("=$A$1*2");

        var map = XlSharedFormulaOptimizer.Build(ws);
        Assert.NotNull(map);
    }

    [Fact]
    public void RoundTrip_SharedFormulasPreserved()
    {
        // Save with shared formulas, load, formulas should be intact
        var path = Path.Combine(Path.GetTempPath(), $"shared-{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = new XlWorkbook();
            var ws = wb.NewWorksheet("Sheet1");
            ws.NewRow(["10"]);
            ws.NewRow(["20"]);
            ws.NewRow(["30"]);
            ws.Rows[0].Insert(1, "").SetFormula("=A1*2", "20");
            ws.Rows[1].Insert(1, "").SetFormula("=A2*2", "40");
            ws.Rows[2].Insert(1, "").SetFormula("=A3*2", "60");

            XlHelper.Save(path, wb);

            var wb2 = XlHelper.Load(path);
            var ws2 = wb2.Worksheets[0];

            Assert.Equal("=A1*2", ws2.Rows[0].Cells[1].Formula);
            Assert.Equal("=A2*2", ws2.Rows[1].Cells[1].Formula);
            Assert.Equal("=A3*2", ws2.Rows[2].Cells[1].Formula);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void RoundTrip_SaveAndReloadXmlHasSharedSyntax()
    {
        var path = Path.Combine(Path.GetTempPath(), $"sharedxml-{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = new XlWorkbook();
            var ws = wb.NewWorksheet("Sheet1");
            ws.NewRow(["10"]);
            ws.NewRow(["20"]);
            ws.NewRow(["30"]);
            ws.Rows[0].Insert(1, "").SetFormula("=A1*2", "20");
            ws.Rows[1].Insert(1, "").SetFormula("=A2*2", "40");
            ws.Rows[2].Insert(1, "").SetFormula("=A3*2", "60");

            XlHelper.Save(path, wb);

            // Verify the XML contains shared formula syntax
            using var archive = System.IO.Compression.ZipFile.OpenRead(path);
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
            Assert.NotNull(entry);
            using var reader = new StreamReader(entry!.Open());
            var xml = reader.ReadToEnd();

            Assert.Contains("t=\"shared\"", xml);
            Assert.Contains("si=\"0\"", xml);
            Assert.Contains("ref=\"B1:B3\"", xml);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Mixed_SomeSharedSomeNot()
    {
        // B1=A1*2, B2=A2*2 (shared), B3=SUM(A1:A3) (not shared)
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        ws.NewRow(["10"]);
        ws.NewRow(["20"]);
        ws.NewRow(["30"]);
        ws.Rows[0].Insert(1, "").SetFormula("=A1*2", "20");
        ws.Rows[1].Insert(1, "").SetFormula("=A2*2", "40");
        ws.Rows[2].Insert(1, "").SetFormula("=SUM(A1:A3)", "60");

        var map = XlSharedFormulaOptimizer.Build(ws);
        Assert.NotNull(map);

        // B1 and B2 should be in the shared map; B3 should not
        Assert.True(map!.ContainsKey((0, 1)));
        Assert.True(map.ContainsKey((1, 1)));
        Assert.False(map.ContainsKey((2, 1)));
    }

    [Fact]
    public void SaveLoad_WithShared_CalculatesCorrectly()
    {
        var path = Path.Combine(Path.GetTempPath(), $"shcalc-{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = new XlWorkbook();
            var ws = wb.NewWorksheet("Sheet1");
            ws.NewRow(["10"]);
            ws.NewRow(["20"]);
            ws.NewRow(["30"]);
            ws.Rows[0].Insert(1, "").SetFormula("=A1*2", "20");
            ws.Rows[1].Insert(1, "").SetFormula("=A2*2", "40");
            ws.Rows[2].Insert(1, "").SetFormula("=A3*2", "60");

            XlHelper.Save(path, wb);

            var wb2 = XlHelper.Load(path);
            var engine = new FormulaEngine();
            engine.Calculate(wb2.Worksheets[0]);

            Assert.Equal("20", wb2.Worksheets[0].Rows[0].Cells[1].Value);
            Assert.Equal("40", wb2.Worksheets[0].Rows[1].Cells[1].Value);
            Assert.Equal("60", wb2.Worksheets[0].Rows[2].Cells[1].Value);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
