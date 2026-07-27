using BenchmarkDotNet.Attributes;

namespace ExcelIO.Benchmark;

[MemoryDiagnoser]
public class CellAccessBenchmarks
{
    private XlWorkbook? _wb;
    private XlWorksheet? _ws;

    [GlobalSetup]
    public void Setup()
    {
        _wb = BenchmarkHelper.CreateWorkbook(500, 200);
        _ws = _wb.Worksheets[0];
        _ws.MapHeaders(new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" });
    }

    [Benchmark]
    public int ForLoop_RowColumn()
    {
        var ws = _ws!;
        int total = 0;
        for (int r = 0; r < ws.Rows.Count; r++)
        {
            var row = ws.Rows[r];
            for (int c = 0; c < row.Cells.Count; c++)
            {
                total += row[c].Length;
            }
        }
        return total;
    }

    [Benchmark]
    public int ForEach_RowsThenCells()
    {
        int total = 0;
        foreach (var row in _ws!.Rows)
        {
            foreach (var cell in row.Cells)
            {
                total += cell.Value.Length;
            }
        }
        return total;
    }

    [Benchmark]
    public int IndexerAccess()
    {
        var ws = _ws!;
        int total = 0;
        for (int r = 0; r < ws.Rows.Count; r++)
        {
            for (int c = 0; c < ws.Rows[r].Cells.Count; c++)
            {
                total += ws.Rows[r][c].Length;
            }
        }
        return total;
    }

    [Benchmark]
    public int ColumnNameAccess()
    {
        var ws = _ws!;
        int total = 0;
        var cols = new[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
        for (int r = 0; r < ws.Rows.Count; r++)
        {
            var row = ws.Rows[r];
            foreach (var c in cols)
            {
                total += row[c].Length;
            }
        }
        return total;
    }

    [Benchmark]
    public int MapHeaders_ThenByName()
    {
        var ws = _ws!;
        int total = 0;
        for (int r = 0; r < ws.Rows.Count; r++)
        {
            total += ws.Cell(r, "A").Value.Length;
        }
        return total;
    }

    [Benchmark]
    public string SetCellValue_10kCells()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        for (int r = 0; r < 100; r++)
        {
            var row = ws.NewRow();
            for (int c = 0; c < 100; c++)
            {
                row.Add("value");
            }
        }
        // Modify half the cells
        for (int r = 0; r < 100; r++)
        {
            for (int c = 0; c < 100; c += 2)
            {
                ws.Rows[r].Cells[c].Value = "updated";
            }
        }
        return ws.Rows[0][0];
    }
}
