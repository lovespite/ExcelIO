using BenchmarkDotNet.Attributes;

namespace ExcelIO.Benchmark;

[MemoryDiagnoser]
public class SaveBenchmarks
{
    private string? _path;

    [GlobalCleanup]
    public void Cleanup()
    {
        if (_path != null && File.Exists(_path))
            File.Delete(_path);
    }

    [Benchmark]
    public void SmallSheet_10kCells()
    {
        var wb = BenchmarkHelper.CreateWorkbook(100, 100);
        _path = BenchmarkHelper.GetTempPath("save-small");
        XlHelper.Save(_path, wb);
    }

    [Benchmark]
    public void MediumSheet_100kCells()
    {
        var wb = BenchmarkHelper.CreateWorkbook(1000, 100);
        _path = BenchmarkHelper.GetTempPath("save-med");
        XlHelper.Save(_path, wb);
    }

    [Benchmark]
    public void WithStyles_10kCells()
    {
        var wb = BenchmarkHelper.CreateStyledWorkbook(100, 100, styleCount: 10);
        _path = BenchmarkHelper.GetTempPath("save-styles");
        XlHelper.Save(_path, wb);
    }

    [Benchmark]
    public void WithMultipleStyles_10kCells()
    {
        var wb = BenchmarkHelper.CreateStyledWorkbook(100, 100, styleCount: 100);
        _path = BenchmarkHelper.GetTempPath("save-multistyle");
        XlHelper.Save(_path, wb);
    }

    [Benchmark]
    public void WithFormulas_10kCells()
    {
        var wb = BenchmarkHelper.CreateFormulaWorkbook(100, 100);
        _path = BenchmarkHelper.GetTempPath("save-formulas");
        XlHelper.Save(_path, wb);
    }

    [Benchmark]
    public void WithImages_10kCells()
    {
        var wb = BenchmarkHelper.CreateWorkbookWithImages(100, 100, imageCount: 10);
        _path = BenchmarkHelper.GetTempPath("save-imgs");
        XlHelper.Save(_path, wb);
    }
}
