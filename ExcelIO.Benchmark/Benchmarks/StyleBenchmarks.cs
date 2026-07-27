using BenchmarkDotNet.Attributes;

namespace ExcelIO.Benchmark;

[MemoryDiagnoser]
public class StyleBenchmarks
{
    private string? _tempPath;

    [GlobalCleanup]
    public void Cleanup()
    {
        if (_tempPath != null && File.Exists(_tempPath))
            File.Delete(_tempPath);
    }

    /// <summary>Save with 1000 cells all sharing 1 style (tests dedup).</summary>
    [Benchmark]
    public void SameStyle_10kCells()
    {
        var wb = BenchmarkHelper.CreateStyledWorkbook(100, 100, styleCount: 1);
        _tempPath = BenchmarkHelper.GetTempPath("style-same");
        XlHelper.Save(_tempPath, wb);
    }

    /// <summary>Save with 10 distinct styles repeated across 10k cells.</summary>
    [Benchmark]
    public void FewStyles_10kCells()
    {
        var wb = BenchmarkHelper.CreateStyledWorkbook(100, 100, styleCount: 10);
        _tempPath = BenchmarkHelper.GetTempPath("style-few");
        XlHelper.Save(_tempPath, wb);
    }

    /// <summary>Save with 100 distinct styles across 10k cells.</summary>
    [Benchmark]
    public void ManyStyles_10kCells()
    {
        var wb = BenchmarkHelper.CreateStyledWorkbook(100, 100, styleCount: 100);
        _tempPath = BenchmarkHelper.GetTempPath("style-many");
        XlHelper.Save(_tempPath, wb);
    }

    /// <summary>Apply bold+color to 100x100 range via XlRange.SetStyle.</summary>
    [Benchmark]
    public void ApplyStyleToRange_10kCells()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        for (int r = 0; r < 100; r++)
            ws.NewRow().AddRange(Enumerable.Range(0, 100).Select(i => $"Cell_{r}_{i}").ToArray());

        var s = new XlStyle { Bold = true, FillColor = "FFFFFF00" };
        ws.Range("A1:CV100").SetStyle(s);
    }

    /// <summary>Apply bold+color to 1000x100 range.</summary>
    [Benchmark]
    public void ApplyStyleToRange_100kCells()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        for (int r = 0; r < 1000; r++)
            ws.NewRow().AddRange(Enumerable.Range(0, 100).Select(i => $"Cell_{r}_{i}").ToArray());

        var s = new XlStyle { Bold = true, FillColor = "FFFFFF00" };
        ws.Range("A1:CV1000").SetStyle(s);
    }

    /// <summary>Set style cell-by-cell (for comparison with range).</summary>
    [Benchmark]
    public void ApplyStyle_CellByCell_10kCells()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        for (int r = 0; r < 100; r++)
        {
            var row = ws.NewRow();
            for (int c = 0; c < 100; c++)
            {
                var cell = row.Insert(c, $"Cell_{r}_{c}");
                cell.Style = new XlStyle { Bold = true, FillColor = "FFFFFF00" };
            }
        }
    }
}
