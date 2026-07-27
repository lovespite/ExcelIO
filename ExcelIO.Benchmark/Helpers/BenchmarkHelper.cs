namespace ExcelIO.Benchmark;

public static class BenchmarkHelper
{
    private static readonly byte[] Png1x1 = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ioAAAAASUVORK5CYII=");

    /// <summary>Create workbook filled with plain text cells.</summary>
    public static XlWorkbook CreateWorkbook(int rowCount, int colCount, string cellPrefix = "Cell")
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        for (int r = 0; r < rowCount; r++)
        {
            var row = ws.NewRow();
            for (int c = 0; c < colCount; c++)
            {
                row.Add($"{cellPrefix}{r}_{c}");
            }
        }
        return wb;
    }

    /// <summary>Create workbook with styles.</summary>
    public static XlWorkbook CreateStyledWorkbook(int rowCount, int colCount, int styleCount = 1)
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var styles = new XlStyle[styleCount];
        for (int i = 0; i < styleCount; i++)
        {
            styles[i] = new XlStyle
            {
                FontName = "Arial",
                FontSize = 10 + i % 10,
                Bold = i % 2 == 0,
                Italic = i % 3 == 0,
                FontColor = $"FF{i % 3 * 5:X2}{i % 7 * 3:X2}{i % 11 * 2:X2}",
                FillColor = $"FF{255 - i % 50:X2}{200 - i % 80:X2}{180 - i % 30:X2}",
                Alignment = new XlAlignment
                {
                    Horizontal = (XlHorizontalAlignment)(i % 5),
                    Vertical = (XlVerticalAlignment)(i % 4)
                }
            };
        }

        for (int r = 0; r < rowCount; r++)
        {
            var row = ws.NewRow();
            for (int c = 0; c < colCount; c++)
            {
                var cell = row.Insert(c, $"Cell_{r}_{c}");
                cell.Style = styles[(r * colCount + c) % styleCount];
            }
        }
        return wb;
    }

    /// <summary>Create workbook with formulas.</summary>
    public static XlWorkbook CreateFormulaWorkbook(int rowCount, int colCount)
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        for (int r = 0; r < rowCount; r++)
        {
            var row = ws.NewRow();
            for (int c = 0; c < colCount; c++)
            {
                if (c > 0)
                {
                    var col = GetColumnName(c);
                    var prevCol = GetColumnName(c - 1);
                    var cell = row.Insert(c, (r * 10).ToString());
                    cell.SetFormula($"={prevCol}{r + 1}*2", (r * 10).ToString());
                }
                else
                {
                    row.Add((r * 5).ToString());
                }
            }
        }
        return wb;
    }

    /// <summary>Create workbook with images.</summary>
    public static XlWorkbook CreateWorkbookWithImages(int rowCount, int colCount, int imageCount)
    {
        var wb = CreateWorkbook(rowCount, colCount);
        var ws = wb.Worksheets[0];
        for (int i = 0; i < imageCount; i++)
        {
            ws.AddImage(Png1x1, "png", rowIndex: i % rowCount, columnIndex: i % colCount);
        }
        return wb;
    }

    public static string GetColumnName(int index)
    {
        string name = "";
        while (index >= 0)
        {
            name = (char)('A' + (index % 26)) + name;
            index = (index / 26) - 1;
        }
        return name;
    }

    /// <summary>Get temp file path and memory baseline.</summary>
    public static string GetTempPath(string prefix) =>
        Path.Combine(Path.GetTempPath(), $"excelio-bm-{prefix}-{Guid.NewGuid():N}.xlsx");

    /// <summary>Force GC and return baseline memory.</summary>
    public static long GetMemoryBaseline()
    {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        return GC.GetTotalMemory(forceFullCollection: false);
    }
}
