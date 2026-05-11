namespace ExcelIO;

public record struct XlPoint(int Column, int Row);

public class XlRange
{
    private readonly XlWorksheet _ws;
    
    public string RangeExpression { get; }
    public int StartColumn { get; }
    public int StartRow { get; }
    public int EndColumn { get; }
    public int EndRow { get; }
    
    public bool IsInfiniteRow { get; }
    public bool IsInfiniteColumn { get; }

    public XlPoint StartPosition => new(StartColumn, StartRow);
    public XlPoint EndPosition => (IsInfiniteRow || IsInfiniteColumn) ? new(-1, -1) : new(EndColumn, EndRow);

    public XlRange(XlWorksheet ws, int col, int row)
    {
        _ws = ws;
        StartColumn = EndColumn = col;
        StartRow = EndRow = row;
        RangeExpression = GetColumnName(col) + (row + 1);
    }

    public XlRange(XlWorksheet ws, string expression)
    {
        _ws = ws;
        RangeExpression = expression;
        
        var parts = expression.Split(':');
        if (parts.Length == 1)
        {
            ParseSingle(parts[0], out var sc, out var sr, out var ec, out var er, out var ir, out var ic);
            StartColumn = sc; StartRow = sr; EndColumn = ec; EndRow = er;
            IsInfiniteRow = ir; IsInfiniteColumn = ic;
        }
        else if (parts.Length == 2)
        {
            ParseSingle(parts[0], out var sc1, out var sr1, out _, out _, out var ir1, out var ic1);
            ParseSingle(parts[1], out var sc2, out var sr2, out _, out _, out var ir2, out var ic2);
            
            StartColumn = Math.Min(sc1, sc2);
            EndColumn = Math.Max(sc1, sc2);
            StartRow = Math.Min(sr1, sr2);
            EndRow = Math.Max(sr1, sr2);
            
            IsInfiniteRow = ir1 || ir2;
            IsInfiniteColumn = ic1 || ic2;
        }
        else
        {
            throw new ArgumentException("Invalid range expression", nameof(expression));
        }
    }

    private static void ParseSingle(string part, out int col, out int row, out int endCol, out int endRow, out bool isInfRow, out bool isInfCol)
    {
        col = row = endCol = endRow = 0;
        isInfRow = isInfCol = false;

        if (int.TryParse(part, out var rowIdx))
        {
            row = endRow = rowIdx - 1;
            isInfCol = true;
            return;
        }

        if (part.All(char.IsLetter))
        {
            col = endCol = GetColumnIndex(part);
            isInfRow = true;
            return;
        }

        // A1 format
        int i = 0;
        while (i < part.Length && char.IsLetter(part[i])) i++;
        var colPart = part[..i];
        var rowPart = part[i..];

        col = endCol = GetColumnIndex(colPart);
        row = endRow = int.Parse(rowPart) - 1;
    }

    public static int GetColumnIndex(string name)
    {
        int index = 0;
        name = name.ToUpperInvariant();
        for (int i = 0; i < name.Length; i++)
        {
            index *= 26;
            index += (name[i] - 'A' + 1);
        }
        return index - 1;
    }

    public static string GetColumnName(int index)
    {
        string columnName = "";
        while (index >= 0)
        {
            columnName = (char)('A' + (index % 26)) + columnName;
            index = (index / 26) - 1;
        }
        return columnName;
    }

    public void SetStyle(XlStyle style)
    {
        if (IsInfiniteRow)
        {
            for (int c = StartColumn; c <= EndColumn; c++)
            {
                if (!_ws.Columns.TryGetValue(c, out var col))
                {
                    col = new XlColumn();
                    _ws.Columns[c] = col;
                }
                col.Style = style;
            }
        }
        else if (IsInfiniteColumn)
        {
            EnsureRows(EndRow);
            for (int r = StartRow; r <= EndRow; r++)
            {
                _ws.Rows[r].Style = style;
            }
        }
        else
        {
            EnsureRows(EndRow);
            for (int r = StartRow; r <= EndRow; r++)
            {
                var row = _ws.Rows[r];
                for (int c = StartColumn; c <= EndColumn; c++)
                {
                    while (row.Cells.Count <= c) row.Cells.Add(new XlCell(row));
                    row.Cells[c].Style = style;
                }
            }
        }
    }

    public void ClearStyle() => SetStyle(null!);

    public void SetContent(string content)
    {
        if (IsInfiniteRow || IsInfiniteColumn) return; // Fail silently

        EnsureRows(EndRow);
        for (int r = StartRow; r <= EndRow; r++)
        {
            var row = _ws.Rows[r];
            for (int c = StartColumn; c <= EndColumn; c++)
            {
                while (row.Cells.Count <= c) row.Cells.Add(new XlCell(row));
                row.Cells[c].Value = content;
            }
        }
    }

    public void ClearContent() => SetContent(string.Empty);

    public void Merge()
    {
        var refStr = GetReferenceString();
        if (!_ws.MergedCells.Contains(refStr))
        {
            _ws.MergedCells.Add(refStr);
        }
    }

    public void Unmerge()
    {
        var refStr = GetReferenceString();
        _ws.MergedCells.Remove(refStr);
    }

    private string GetReferenceString()
    {
        if (IsInfiniteRow) return $"{GetColumnName(StartColumn)}:{GetColumnName(EndColumn)}";
        if (IsInfiniteColumn) return $"{StartRow + 1}:{EndRow + 1}";
        if (StartColumn == EndColumn && StartRow == EndRow) return $"{GetColumnName(StartColumn)}{StartRow + 1}";
        return $"{GetColumnName(StartColumn)}{StartRow + 1}:{GetColumnName(EndColumn)}{EndRow + 1}";
    }

    private void EnsureRows(int maxRow)
    {
        while (_ws.Rows.Count <= maxRow)
        {
            _ws.Rows.Add(new XlRow(_ws));
        }
    }
}
