using System.IO;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Text;

namespace ExcelIO;

public static partial class ExcelIOWrapper
{
    private static XlWorkbook? _workbook;
    private static int _activeSheetIndex;
    private static string _fileName = "workbook.xlsx";

    private static XlWorksheet? ActiveSheet =>
        _workbook?.Worksheets.ElementAtOrDefault(_activeSheetIndex);

    // ── Workbook operations ──

    [JSExport]
    public static bool LoadFromBytes(byte[] data, string fileName)
    {
        try
        {
            using var ms = new MemoryStream(data);
            _workbook = XlHelper.Load(ms);
            _activeSheetIndex = 0;
            _fileName = Path.GetFileNameWithoutExtension(fileName) + ".xlsx";
            return true;
        }
        catch
        {
            return false;
        }
    }

    [JSExport]
    public static byte[] SaveToBytes()
    {
        if (_workbook is null)
            return [];

        using var ms = new MemoryStream();
        XlHelper.Save(ms, _workbook);
        return ms.ToArray();
    }

    [JSExport]
    public static void NewWorkbook()
    {
        _workbook = new XlWorkbook();
        _workbook.NewWorksheet("Sheet1");
        _activeSheetIndex = 0;
        _fileName = "workbook.xlsx";
    }

    [JSExport]
    public static int GetSheetCount() =>
        _workbook?.Worksheets.Count ?? 0;

    [JSExport]
    public static string GetSheetName(int index) =>
        _workbook?.Worksheets.ElementAtOrDefault(index)?.Name ?? "";

    [JSExport]
    public static int GetActiveSheetIndex() =>
        _activeSheetIndex;

    [JSExport]
    public static bool SwitchSheet(int index)
    {
        if (_workbook is null || index < 0 || index >= _workbook.Worksheets.Count)
            return false;
        _activeSheetIndex = index;
        return true;
    }

    [JSExport]
    public static string GetFileName() => _fileName;

    // ── Data queries ──

    [JSExport]
    public static int GetRowCount() =>
        ActiveSheet?.Rows.Count ?? 0;

    [JSExport]
    public static int GetColCount()
    {
        var sheet = ActiveSheet;
        if (sheet is null) return 0;
        int maxCols = 0;
        foreach (var row in sheet.Rows)
        {
            if (row.Cells.Count > maxCols)
                maxCols = row.Cells.Count;
        }
        return maxCols;
    }

    [JSExport]
    public static string GetCellValue(int row, int col)
    {
        var sheet = ActiveSheet;
        if (sheet is null || row < 0 || row >= sheet.Rows.Count)
            return "";
        var r = sheet.Rows[row];
        if (col < 0 || col >= r.Cells.Count)
            return "";
        return r.Cells[col].Value;
    }

    [JSExport]
    public static void SetCellValue(int row, int col, string value)
    {
        var sheet = ActiveSheet;
        if (sheet is null)
            return;

        while (sheet.Rows.Count <= row)
            sheet.Rows.Add(new XlRow(sheet));

        var r = sheet.Rows[row];
        while (r.Cells.Count <= col)
            r.Cells.Add(new XlCell(r));

        r.Cells[col].Value = value;
    }

    [JSExport]
    public static void AddRow(string[] values)
    {
        ActiveSheet?.AddRow(values);
    }

    [JSExport]
    public static void ClearRows()
    {
        ActiveSheet?.ClearRows();
    }

    // ── Style queries (return JSON) ──

    [JSExport]
    public static string GetCellStyleJson(int row, int col)
    {
        var sheet = ActiveSheet;
        if (sheet is null || row < 0 || row >= sheet.Rows.Count)
            return "{}";
        var r = sheet.Rows[row];
        if (col < 0 || col >= r.Cells.Count)
            return "{}";

        var cellStyle = r.Cells[col].Style;

        // Resolve cascaded style
        var rowStyle = r.Style;
        XlStyle? colStyle = null;
        if (sheet.Columns.TryGetValue(col, out var colObj))
            colStyle = colObj.Style;

        return StyleToJson(MergeStyles(colStyle, rowStyle, cellStyle));
    }

    [JSExport]
    public static string GetRowStyleJson(int row)
    {
        var sheet = ActiveSheet;
        if (sheet is null || row < 0 || row >= sheet.Rows.Count)
            return "{}";
        return StyleToJson(sheet.Rows[row].Style);
    }

    [JSExport]
    public static string GetColStyleJson(int col)
    {
        var sheet = ActiveSheet;
        if (sheet is null || !sheet.Columns.TryGetValue(col, out var colObj) || colObj?.Style is null)
            return "{}";
        return StyleToJson(colObj.Style);
    }

    [JSExport]
    public static string GetMergedCells()
    {
        var sheet = ActiveSheet;
        if (sheet is null || sheet.MergedCells.Count == 0)
            return "[]";

        var sb = new StringBuilder();
        sb.Append('[');
        for (int i = 0; i < sheet.MergedCells.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append('"');
            sb.Append(sheet.MergedCells[i]);
            sb.Append('"');
        }
        sb.Append(']');
        return sb.ToString();
    }

    [JSExport]
    public static double GetColumnWidth(int col)
    {
        var sheet = ActiveSheet;
        if (sheet is null || !sheet.Columns.TryGetValue(col, out var colObj) || colObj?.Width is null)
            return 0;
        return colObj.Width.Value;
    }

    [JSExport]
    public static double GetRowHeight(int row)
    {
        var sheet = ActiveSheet;
        if (sheet is null || row < 0 || row >= sheet.Rows.Count || sheet.Rows[row].Height is null)
            return 0;
        return sheet.Rows[row].Height!.Value;
    }

    [JSExport]
    public static bool IsRowHidden(int row)
    {
        var sheet = ActiveSheet;
        if (sheet is null || row < 0 || row >= sheet.Rows.Count)
            return false;
        return sheet.Rows[row].Hidden;
    }

    [JSExport]
    public static bool IsColumnHidden(int col)
    {
        var sheet = ActiveSheet;
        if (sheet is null || !sheet.Columns.TryGetValue(col, out var colObj) || colObj is null)
            return false;
        return colObj.Hidden;
    }

    [JSExport]
    public static string GetSheetTabColor()
    {
        var sheet = ActiveSheet;
        if (sheet?.Options.TabColor is null)
            return "";
        return ArgbToCssColor(sheet.Options.TabColor);
    }

    // ── Helpers ──

    private static string StyleToJson(XlStyle? style)
    {
        if (style is null) return "{}";

        var sb = new StringBuilder();
        sb.Append('{');

        AppendJsonProperty(sb, "fontName", style.FontName);
        if (style.FontSize.HasValue)
            AppendJsonNumber(sb, "fontSize", style.FontSize.Value);
        if (style.FontColor is not null)
            AppendJsonProperty(sb, "fontColor", ArgbToCssColor(style.FontColor));
        if (style.Bold)
            AppendJsonBool(sb, "bold", true);
        if (style.Italic)
            AppendJsonBool(sb, "italic", true);
        if (style.FillColor is not null)
            AppendJsonProperty(sb, "fillColor", ArgbToCssColor(style.FillColor));

        if (style.Alignment is not null)
        {
            AppendJsonProperty(sb, "hAlign", HAlignToString(style.Alignment.Horizontal));
            AppendJsonProperty(sb, "vAlign", VAlignToString(style.Alignment.Vertical));
            if (style.Alignment.WrapText)
                AppendJsonBool(sb, "wrapText", true);
        }

        if (style.Border is not null)
        {
            AppendJsonProperty(sb, "borderLeft", BorderStyleToString(style.Border.Left));
            if (style.Border.LeftColor is not null)
                AppendJsonProperty(sb, "borderLeftColor", ArgbToCssColor(style.Border.LeftColor));
            AppendJsonProperty(sb, "borderRight", BorderStyleToString(style.Border.Right));
            if (style.Border.RightColor is not null)
                AppendJsonProperty(sb, "borderRightColor", ArgbToCssColor(style.Border.RightColor));
            AppendJsonProperty(sb, "borderTop", BorderStyleToString(style.Border.Top));
            if (style.Border.TopColor is not null)
                AppendJsonProperty(sb, "borderTopColor", ArgbToCssColor(style.Border.TopColor));
            AppendJsonProperty(sb, "borderBottom", BorderStyleToString(style.Border.Bottom));
            if (style.Border.BottomColor is not null)
                AppendJsonProperty(sb, "borderBottomColor", ArgbToCssColor(style.Border.BottomColor));
        }

        // Remove trailing comma
        if (sb[^1] == ',')
            sb.Length--;

        sb.Append('}');
        return sb.ToString();
    }

    private static void AppendJsonProperty(StringBuilder sb, string key, string? value)
    {
        if (value is null) return;
        sb.Append('"');
        sb.Append(key);
        sb.Append("\":\"");
        sb.Append(EscapeJson(value));
        sb.Append("\",");
    }

    private static void AppendJsonNumber(StringBuilder sb, string key, double value)
    {
        sb.Append('"');
        sb.Append(key);
        sb.Append("\":");
        sb.Append(value.ToString("0.##"));
        sb.Append(',');
    }

    private static void AppendJsonBool(StringBuilder sb, string key, bool value)
    {
        sb.Append('"');
        sb.Append(key);
        sb.Append("\":");
        sb.Append(value ? "true" : "false");
        sb.Append(',');
    }

    private static string EscapeJson(string s)
    {
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private static string ArgbToCssColor(string argb)
    {
        // "FFFF0000" → "#FF0000", "FF0000" → "#FF0000"
        if (argb.Length >= 8)
            return "#" + argb.Substring(2, 6);
        if (argb.Length == 6)
            return "#" + argb;
        return "#" + argb;
    }

    private static string? HAlignToString(XlHorizontalAlignment h) => h switch
    {
        XlHorizontalAlignment.Left => "left",
        XlHorizontalAlignment.Center => "center",
        XlHorizontalAlignment.Right => "right",
        XlHorizontalAlignment.Justify => "justify",
        _ => null
    };

    private static string? VAlignToString(XlVerticalAlignment v) => v switch
    {
        XlVerticalAlignment.Top => "top",
        XlVerticalAlignment.Center => "middle",
        XlVerticalAlignment.Bottom => "bottom",
        XlVerticalAlignment.Justify => "justify",
        _ => null
    };

    private static string? BorderStyleToString(XlBorderStyle s) => s switch
    {
        XlBorderStyle.None => null,
        XlBorderStyle.Thin => "thin",
        XlBorderStyle.Medium => "medium",
        XlBorderStyle.Dashed => "dashed",
        XlBorderStyle.Dotted => "dotted",
        XlBorderStyle.Thick => "thick",
        XlBorderStyle.Double => "double",
        XlBorderStyle.Hair => "hair",
        _ => null
    };

    /// <summary>
    /// Merge styles with cascade: column (lowest) → row → cell (highest priority)
    /// </summary>
    private static XlStyle? MergeStyles(XlStyle? colStyle, XlStyle? rowStyle, XlStyle? cellStyle)
    {
        if (cellStyle is not null) return cellStyle;
        if (rowStyle is not null) return rowStyle;
        return colStyle;
    }
}
