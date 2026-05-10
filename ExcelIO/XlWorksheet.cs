using System.Collections;

namespace ExcelIO;

public class XlWorksheet : IReadOnlyList<XlRow>
{
    private readonly Dictionary<string, int> _indexCache = new(StringComparer.OrdinalIgnoreCase);

    private readonly XlWorkbook _wb;

    public XlWorkbook Workbook => _wb;

    public XlWorksheet(XlWorkbook wb)
    {
        _wb = wb;
    }

    public int IndexOf(string columnName)
    {
        if (_indexCache.TryGetValue(columnName, out int index))
        {
            return index;
        }
        return -1;
    }

    /// <summary>
    /// Copies a range of rows from the current worksheet to the specified destination worksheet.
    /// </summary>
    /// <remarks>Rows are copied in order and inserted into the destination worksheet starting at <paramref
    /// name="targetIndex"/> if specified; otherwise, they are appended. The copied rows contain the cell values from
    /// the source rows, but do not preserve references to the original row or worksheet objects.</remarks>
    /// <param name="start">The zero-based index of the first row to copy from the source worksheet. Must be greater than or equal to 0 and
    /// less than the total number of rows.</param>
    /// <param name="count">The number of rows to copy, starting from <paramref name="start"/>. Must be greater than 0, and the range
    /// defined by <paramref name="start"/> and <paramref name="count"/> must not exceed the bounds of the source
    /// worksheet.</param>
    /// <param name="dest">The destination worksheet to which the rows will be copied. Cannot be null.</param>
    /// <param name="targetIndex">The zero-based index in the destination worksheet at which to insert the copied rows. If less than 0 or greater
    /// than the number of rows in the destination worksheet, rows are appended to the end.</param>
    /// <exception cref="IndexOutOfRangeException">Thrown when the specified source range, defined by <paramref name="start"/> and <paramref name="count"/>, is
    /// outside the bounds of the source worksheet.</exception>
    public IReadOnlyList<XlRow> CopyRowsTo(int start, int count, XlWorksheet dest, int targetIndex = -1)
    {
        var endIndex = start + count - 1;

        if (start < 0 || endIndex >= Rows.Count)
        {
            throw new IndexOutOfRangeException($"Source range is out of bounds. Total rows: {Rows.Count}");
        }

        var rows = new List<XlRow>(count);

        for (var rIndex = start; rIndex <= endIndex; rIndex++)
        {
            var sourceRow = Rows[rIndex];
            var destRow = new XlRow(dest);
            destRow.AddRange(sourceRow.Cells.Select(c => c.Value));
            if (targetIndex >= 0 && targetIndex < dest.Rows.Count)
            {
                dest.Rows.Insert(targetIndex, destRow);
                targetIndex++;
            }
            else
            {
                dest.Rows.Add(destRow);
            }
            rows.Add(destRow);
        }

        return rows;
    }

    public void MapHeaders(IReadOnlyList<string> headers)
    {
        _indexCache.Clear();
        if (Rows.Count == 0)
        {
            AddRow([.. headers]);
            for (var i = 0; i < headers.Count; i++)
            {
                _indexCache[headers[i]] = i;
            }
            return;
        }

        List<string> headerList = [.. Rows[0]];

        for (int i = 0; i < headers.Count; i++)
        {
            var headerValue = headers[i];
            var index = FindIndex(headerValue, headerList);
            if (index < 0) continue;
            _indexCache[headerValue] = index;
        }
    }

    public XlCell Cell(int rowIndex, string columnName)
    {
        if (rowIndex < 0 || rowIndex >= Rows.Count)
        {
            throw new IndexOutOfRangeException($"行索引 {rowIndex} 超出了范围， 最大行数：{Rows.Count}");
        }
        if (!_indexCache.TryGetValue(columnName, out int colIndex))
        {
            colIndex = FindIndex(columnName, [.. Rows[0].Cells.Select(c => c.Value)]);
            if (colIndex < 0) throw new KeyNotFoundException($"列 '{columnName}' 不存在");
            else _indexCache[columnName] = colIndex;
        }

        if (Rows[rowIndex].Cells.Count <= colIndex) return XlCell.Null;
        return Rows[rowIndex].Cells[colIndex];
    }

    private static int FindIndex(string name, List<string> list)
    {
        for (int i = 0; i < list.Count; i++)
        {
            if (string.Equals(list[i], name, StringComparison.OrdinalIgnoreCase))
                return i;
        }
        return -1;
    }

    public XlRow this[int index]
    {
        get
        {
            return Rows[index];
        }
    }

    public string Name { get; set; } = "Sheet1";
    public List<XlRow> Rows { get; set; } = [];
    public List<XlWorksheetImage> Images { get; } = [];

    public int Count => ((IReadOnlyCollection<XlRow>)Rows).Count;

    // 辅助方法：快速添加一行
    public void AddRow(params string[] values)
    {
        var row = new XlRow(this);
        row.AddRange(values);
        Rows.AddRange(row);
    }

    public void AddRow(IEnumerable<string> values)
    {
        var row = new XlRow(this);
        row.AddRange(values);
        Rows.AddRange(row);
    }

    public void AddRows(List<string[]> rows)
    {
       _ = rows.Select(NewRow).ToArray();
    }

    public XlRow NewRow(string[]? values = null)
    {
        var row = new XlRow(this);
        Rows.Add(row);
        if (values is not null) row.AddRange(values);
        return row;
    }

    public void ClearRows()
    {
        Rows.Clear();
    }

    public XlWorksheetImage AddImage(string imagePath, int rowIndex, int columnIndex, int rowSpan = 1, int columnSpan = 1)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(imagePath);
        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException("Image file not found.", imagePath);
        }

        var imageBytes = File.ReadAllBytes(imagePath);
        var imageExtension = Path.GetExtension(imagePath);
        return AddImage(imageBytes, imageExtension, rowIndex, columnIndex, rowSpan, columnSpan);
    }

    public XlWorksheetImage AddImage(byte[] imageBytes, string imageExtension, int rowIndex, int columnIndex, int rowSpan = 1, int columnSpan = 1)
    {
        var image = XlWorksheetImage.Create(imageBytes, imageExtension, rowIndex, columnIndex, rowSpan, columnSpan);
        Images.Add(image);
        return image;
    }


    public IEnumerator<XlRow> GetEnumerator()
    {
        return ((IEnumerable<XlRow>)Rows).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable)Rows).GetEnumerator();
    }
}
