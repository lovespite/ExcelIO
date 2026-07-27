using BenchmarkDotNet.Attributes;

namespace ExcelIO.Benchmark;

[MemoryDiagnoser]
public class LoadBenchmarks
{
    private string? _smallPath;
    private string? _mediumPath;
    private string? _stylesPath;
    private string? _formulasPath;
    private string? _imagesPath;

    [GlobalSetup]
    public void Setup()
    {
        var smallWb = BenchmarkHelper.CreateWorkbook(100, 100);
        _smallPath = BenchmarkHelper.GetTempPath("load-small");
        XlHelper.Save(_smallPath, smallWb);

        var medWb = BenchmarkHelper.CreateWorkbook(1000, 100);
        _mediumPath = BenchmarkHelper.GetTempPath("load-med");
        XlHelper.Save(_mediumPath, medWb);

        var styledWb = BenchmarkHelper.CreateStyledWorkbook(100, 100, styleCount: 10);
        _stylesPath = BenchmarkHelper.GetTempPath("load-styles");
        XlHelper.Save(_stylesPath, styledWb);

        var formulaWb = BenchmarkHelper.CreateFormulaWorkbook(100, 100);
        _formulasPath = BenchmarkHelper.GetTempPath("load-formulas");
        XlHelper.Save(_formulasPath, formulaWb);

        var imgWb = BenchmarkHelper.CreateWorkbookWithImages(100, 100, imageCount: 10);
        _imagesPath = BenchmarkHelper.GetTempPath("load-imgs");
        XlHelper.Save(_imagesPath, imgWb);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        foreach (var p in new[] { _smallPath, _mediumPath, _stylesPath, _formulasPath, _imagesPath })
        {
            if (p != null && File.Exists(p)) File.Delete(p);
        }
    }

    [Benchmark]
    public XlWorkbook SmallSheet_10kCells()
    {
        var wb = XlHelper.Load(_smallPath!);
        ForEachCell(wb);
        return wb;
    }

    [Benchmark]
    public XlWorkbook MediumSheet_100kCells()
    {
        var wb = XlHelper.Load(_mediumPath!);
        ForEachCell(wb);
        return wb;
    }

    [Benchmark]
    public XlWorkbook WithStyles_10kCells()
    {
        var wb = XlHelper.Load(_stylesPath!);
        ForEachCell(wb);
        return wb;
    }

    [Benchmark]
    public XlWorkbook WithFormulas_10kCells()
    {
        var wb = XlHelper.Load(_formulasPath!);
        ForEachCell(wb);
        return wb;
    }

    [Benchmark]
    public XlWorkbook WithImages_10kCells()
    {
        var wb = XlHelper.Load(_imagesPath!);
        ForEachCell(wb);
        return wb;
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
