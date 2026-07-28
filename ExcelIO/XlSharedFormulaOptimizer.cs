namespace ExcelIO;

/// <summary>
/// Entry in the shared formula map describing how a cell's formula should be serialized.
/// </summary>
public sealed class SharedFormulaEntry
{
    public int Si { get; }
    public bool IsMaster { get; }
    public string? Ref { get; }
    public string? Formula { get; }

    public SharedFormulaEntry(int si, bool isMaster, string? rangeRef, string? formula)
    {
        Si = si;
        IsMaster = isMaster;
        Ref = rangeRef;
        Formula = formula;
    }
}

/// <summary>
/// Detects groups of formula cells sharing the same pattern so GenerateSheetXml
/// can emit OpenXML shared-formula syntax, reducing file size.
/// </summary>
public static class XlSharedFormulaOptimizer
{
    public static Dictionary<(int Row, int Col), SharedFormulaEntry>? Build(XlWorksheet sheet)
    {
        var map = new Dictionary<(int, int), SharedFormulaEntry>();
        var grouped = new HashSet<(int, int)>();
        int si = 0;

        int maxCol = 0;
        foreach (var row in sheet.Rows)
            if (row.Cells.Count > maxCol) maxCol = row.Cells.Count;

        // ── Vertical sharing (same column, consecutive rows) ──
        for (int c = 0; c < maxCol; c++)
        {
            int r = 0;
            while (r < sheet.Rows.Count)
            {
                while (r < sheet.Rows.Count && !HasFormula(sheet, r, c))
                    r++;
                if (r >= sheet.Rows.Count) break;

                r = TryBuildVerticalGroup(sheet, c, r, map, grouped, ref si);
            }
        }

        // ── Horizontal sharing (same row, consecutive columns) ──
        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            int c = 0;
            while (c < row.Cells.Count)
            {
                while (c < row.Cells.Count && (!row.Cells[c].HasFormula || grouped.Contains((r, c))))
                    c++;
                if (c >= row.Cells.Count) break;

                c = TryBuildHorizontalGroup(sheet, r, c, map, grouped, ref si);
            }
        }

        return map.Count > 0 ? map : null;
    }

    /// <summary>
    /// Try to build a vertical shared-formula group starting at (startRow, col).
    /// Returns the row index to resume scanning from.
    /// </summary>
    private static int TryBuildVerticalGroup(XlWorksheet sheet, int col, int startRow,
        Dictionary<(int, int), SharedFormulaEntry> map, HashSet<(int, int)> grouped, ref int si)
    {
        var masterFormula = sheet.Rows[startRow].Cells[col].Formula!;
        int endRow = startRow;
        while (endRow + 1 < sheet.Rows.Count && HasFormula(sheet, endRow + 1, col))
            endRow++;

        // Collect shareable cells
        var members = new List<(int, int)> { (startRow, col) };
        for (int row = startRow + 1; row <= endRow; row++)
        {
            if (grouped.Contains((row, col))) continue;
            var translated = XlFormulaUtil.TranslateSharedFormula(masterFormula, row - startRow, 0);
            if (string.Equals(translated, sheet.Rows[row].Cells[col].Formula, StringComparison.Ordinal))
                members.Add((row, col));
            else
                break; // pattern broke — don't collect beyond here
        }

        if (members.Count >= 2)
        {
            var refStr = XlFormulaUtil.GetColumnName(col) + (startRow + 1) + ":" +
                         XlFormulaUtil.GetColumnName(col) + (members[^1].Item1 + 1);

            map[(startRow, col)] = new SharedFormulaEntry(si, true, refStr, masterFormula);
            grouped.Add((startRow, col));

            for (int i = 1; i < members.Count; i++)
            {
                map[members[i]] = new SharedFormulaEntry(si, false, null, null);
                grouped.Add(members[i]);
            }
            si++;
        }

        return endRow + 1; // skip the entire run
    }

    /// <summary>
    /// Try to build a horizontal shared-formula group starting at (row, startCol).
    /// Returns the column index to resume scanning from.
    /// </summary>
    private static int TryBuildHorizontalGroup(XlWorksheet sheet, int row, int startCol,
        Dictionary<(int, int), SharedFormulaEntry> map, HashSet<(int, int)> grouped, ref int si)
    {
        var cells = sheet.Rows[row].Cells;
        var masterFormula = cells[startCol].Formula!;

        int endCol = startCol;
        while (endCol + 1 < cells.Count && cells[endCol + 1].HasFormula)
            endCol++;

        var members = new List<(int, int)> { (row, startCol) };
        for (int c = startCol + 1; c <= endCol; c++)
        {
            if (grouped.Contains((row, c))) continue;
            var translated = XlFormulaUtil.TranslateSharedFormula(masterFormula, 0, c - startCol);
            if (string.Equals(translated, cells[c].Formula, StringComparison.Ordinal))
                members.Add((row, c));
            else
                break;
        }

        if (members.Count >= 2)
        {
            var refStr = XlFormulaUtil.GetColumnName(startCol) + (row + 1) + ":" +
                         XlFormulaUtil.GetColumnName(members[^1].Item2) + (row + 1);

            map[(row, startCol)] = new SharedFormulaEntry(si, true, refStr, masterFormula);
            grouped.Add((row, startCol));

            for (int i = 1; i < members.Count; i++)
            {
                map[members[i]] = new SharedFormulaEntry(si, false, null, null);
                grouped.Add(members[i]);
            }
            si++;
        }

        return endCol + 1;
    }

    private static bool HasFormula(XlWorksheet sheet, int row, int col)
    {
        if (row >= sheet.Rows.Count) return false;
        var r = sheet.Rows[row];
        if (col >= r.Cells.Count) return false;
        return r.Cells[col].HasFormula;
    }
}
