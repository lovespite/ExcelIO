using BenchmarkDotNet.Attributes;

namespace ExcelIO.Benchmark;

[MemoryDiagnoser]
public class MemoryBenchmarks
{
    private string? _smallPath;
    private string? _mediumPath;
    private string? _largePath;
    private string? _stylesPath;
    private string? _formulasPath;

    [GlobalSetup]
    public void Setup()
    {
        var smallWb = BenchmarkHelper.CreateWorkbook(100, 100);
        _smallPath = BenchmarkHelper.GetTempPath("mem-small");
        XlHelper.Save(_smallPath, smallWb);

        var medWb = BenchmarkHelper.CreateWorkbook(1000, 100);
        _mediumPath = BenchmarkHelper.GetTempPath("mem-med");
        XlHelper.Save(_mediumPath, medWb);

        var largeWb = BenchmarkHelper.CreateWorkbook(5000, 100);
        _largePath = BenchmarkHelper.GetTempPath("mem-large");
        XlHelper.Save(_largePath, largeWb);

        var styledWb = BenchmarkHelper.CreateStyledWorkbook(100, 100, styleCount: 100);
        _stylesPath = BenchmarkHelper.GetTempPath("mem-styles");
        XlHelper.Save(_stylesPath, styledWb);

        var formulaWb = BenchmarkHelper.CreateFormulaWorkbook(100, 100);
        _formulasPath = BenchmarkHelper.GetTempPath("mem-formulas");
        XlHelper.Save(_formulasPath, formulaWb);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        foreach (var p in new[] { _smallPath, _mediumPath, _largePath, _stylesPath, _formulasPath })
        {
            if (p != null && File.Exists(p)) File.Delete(p);
        }
    }

    // ── Steady-state memory ──

    [Benchmark]
    public long LoadMemory_10kCells()
    {
        BenchmarkHelper.GetMemoryBaseline();
        var wb = XlHelper.Load(_smallPath!);
        ForEachCell(wb);
        return GC.GetTotalMemory(forceFullCollection: true);
    }

    [Benchmark]
    public long LoadMemory_100kCells()
    {
        BenchmarkHelper.GetMemoryBaseline();
        var wb = XlHelper.Load(_mediumPath!);
        ForEachCell(wb);
        return GC.GetTotalMemory(forceFullCollection: true);
    }

    [Benchmark]
    public long LoadMemory_500kCells()
    {
        BenchmarkHelper.GetMemoryBaseline();
        var wb = XlHelper.Load(_largePath!);
        ForEachCell(wb);
        return GC.GetTotalMemory(forceFullCollection: true);
    }

    [Benchmark]
    public long LoadMemory_WithStyles_10k()
    {
        BenchmarkHelper.GetMemoryBaseline();
        var wb = XlHelper.Load(_stylesPath!);
        ForEachCell(wb);
        return GC.GetTotalMemory(forceFullCollection: true);
    }

    [Benchmark]
    public long LoadMemory_WithFormulas_10k()
    {
        BenchmarkHelper.GetMemoryBaseline();
        var wb = XlHelper.Load(_formulasPath!);
        ForEachCell(wb);
        return GC.GetTotalMemory(forceFullCollection: true);
    }

    // ── Memory growth during save ──

    [Benchmark]
    public long SaveWorkbook_MemoryGrowth()
    {
        var wb = BenchmarkHelper.CreateWorkbook(5000, 100);

        BenchmarkHelper.GetMemoryBaseline();
        var before = GC.GetTotalMemory(forceFullCollection: true);

        var path = BenchmarkHelper.GetTempPath("mem-save");
        XlHelper.Save(path, wb);
        var after = GC.GetTotalMemory(forceFullCollection: true);

        if (File.Exists(path)) File.Delete(path);
        return after - before;
    }

    // ── Memory-to-file inflation ratio ──

    [Benchmark]
    public double InflationRatio_10kCells()
    {
        BenchmarkHelper.GetMemoryBaseline();
        var wb = XlHelper.Load(_smallPath!);
        ForEachCell(wb);
        var mem = GC.GetTotalMemory(forceFullCollection: true);
        var fileSize = new FileInfo(_smallPath!).Length;
        return (double)mem / fileSize;
    }

    [Benchmark]
    public double InflationRatio_100kCells()
    {
        BenchmarkHelper.GetMemoryBaseline();
        var wb = XlHelper.Load(_mediumPath!);
        ForEachCell(wb);
        var mem = GC.GetTotalMemory(forceFullCollection: true);
        var fileSize = new FileInfo(_mediumPath!).Length;
        return (double)mem / fileSize;
    }

    [Benchmark]
    public double InflationRatio_500kCells()
    {
        BenchmarkHelper.GetMemoryBaseline();
        var wb = XlHelper.Load(_largePath!);
        ForEachCell(wb);
        var mem = GC.GetTotalMemory(forceFullCollection: true);
        var fileSize = new FileInfo(_largePath!).Length;
        return (double)mem / fileSize;
    }

    private static void ForEachCell(XlWorkbook wb)
    {
        foreach (var ws in wb.Worksheets)
        {
            foreach (var row in ws.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    _ = cell.Value;
                }
            }
        }
    }
}
