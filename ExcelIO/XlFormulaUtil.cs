using System.Text.RegularExpressions;

namespace ExcelIO;

public static class XlFormulaUtil
{
    private static readonly string[] ColumnNameCache = BuildColumnCache();
    private static readonly Regex CellRefPattern = new(@"(\$?)([A-Z]+)(\$?)(\d+)", RegexOptions.Compiled);

    private static string[] BuildColumnCache()
    {
        var cache = new string[1024];
        for (int i = 0; i < 1024; i++)
            cache[i] = GetColumnNameSlow(i);
        return cache;
    }

    public static string GetColumnName(int index)
    {
        if (index >= 0 && index < ColumnNameCache.Length)
            return ColumnNameCache[index];
        return GetColumnNameSlow(index);
    }

    private static string GetColumnNameSlow(int index)
    {
        var columnName = "";
        while (index >= 0)
        {
            columnName = (char)('A' + (index % 26)) + columnName;
            index = (index / 26) - 1;
        }
        return columnName;
    }

    public static string EscapeXml(string txt)
    {
        if (string.IsNullOrEmpty(txt)) return "";
        return txt.Replace("&", "&amp;")
                  .Replace("<", "&lt;")
                  .Replace(">", "&gt;")
                  .Replace("\"", "&quot;")
                  .Replace("'", "&apos;");
    }

    public static (int Row, int Col) ParseCellRef(string cellRef)
    {
        int colEnd = 0;
        while (colEnd < cellRef.Length && char.IsLetter(cellRef[colEnd]))
            colEnd++;

        var colPart = cellRef.Substring(0, colEnd);
        var rowPart = cellRef.Substring(colEnd);

        int col = 0;
        foreach (char c in colPart)
            col = col * 26 + (c - 'A' + 1);
        col -= 1;

        int row = int.Parse(rowPart) - 1;
        return (row, col);
    }

    public static string TranslateSharedFormula(string masterFormula, int rowOffset, int colOffset)
    {
        if (rowOffset == 0 && colOffset == 0)
            return masterFormula;

        return CellRefPattern.Replace(masterFormula, match =>
        {
            var colAbs = match.Groups[1].Value;
            var colPart = match.Groups[2].Value;
            var rowAbs = match.Groups[3].Value;
            var rowPart = match.Groups[4].Value;

            int col = 0;
            foreach (char c in colPart)
                col = col * 26 + (c - 'A' + 1);
            col -= 1;

            int row = int.Parse(rowPart) - 1;

            if (string.IsNullOrEmpty(colAbs))
                col += colOffset;
            if (string.IsNullOrEmpty(rowAbs))
                row += rowOffset;

            return colAbs + GetColumnName(col) + rowAbs + (row + 1).ToString();
        });
    }
}
