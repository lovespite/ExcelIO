using System.IO;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Text;
using System.Text.Json;
using ExcelIO.Formula;

namespace ExcelIO;

public static partial class ExcelIOWrapper
{
    private static XlWorkbook? _workbook;
    private static int _activeSheetIndex;
    private static string _fileName = "workbook.xlsx";
    private static FormulaEngine? _engine;

    private static XlWorksheet? ActiveSheet =>
        _workbook?.Worksheets.ElementAtOrDefault(_activeSheetIndex);

    private static FormulaEngine Engine =>
        _engine ??= new FormulaEngine();

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
    public static string GetCellFormula(int row, int col)
    {
        var sheet = ActiveSheet;
        if (sheet is null || row < 0 || row >= sheet.Rows.Count)
            return "";
        var r = sheet.Rows[row];
        if (col < 0 || col >= r.Cells.Count)
            return "";
        return r.Cells[col].Formula ?? "";
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

    // ── Formula ──

    [JSExport]
    public static void SetCellFormula(int row, int col, string formula)
    {
        var sheet = ActiveSheet;
        if (sheet is null) return;

        while (sheet.Rows.Count <= row)
            sheet.Rows.Add(new XlRow(sheet));

        var r = sheet.Rows[row];
        while (r.Cells.Count <= col)
            r.Cells.Add(new XlCell(r));

        r.Cells[col].SetFormula(formula);
    }

    [JSExport]
    public static void CalculateSheet()
    {
        var sheet = ActiveSheet;
        if (sheet is null) return;
        Engine.Calculate(sheet);
    }

    [JSExport]
    public static void CalculateWorkbook()
    {
        if (_workbook is null) return;
        Engine.Calculate(_workbook);
    }

    [JSExport]
    public static string GetCircularReferences()
    {
        var refs = Engine.CircularReferences;
        if (refs.Count == 0) return "[]";

        var sb = new StringBuilder();
        sb.Append('[');
        for (int i = 0; i < refs.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append("{\"path\":[");
            var path = refs[i].Path;
            for (int j = 0; j < path.Count; j++)
            {
                if (j > 0) sb.Append(',');
                sb.Append("{\"row\":");
                sb.Append(path[j].Row);
                sb.Append(",\"col\":");
                sb.Append(path[j].Col);
                sb.Append('}');
            }
            sb.Append("]}");
        }
        sb.Append(']');
        return sb.ToString();
    }

    // ── Style setters ──

    [JSExport]
    public static void SetCellStyle(int row, int col, string styleJson)
    {
        var sheet = ActiveSheet;
        if (sheet is null) return;

        while (sheet.Rows.Count <= row)
            sheet.Rows.Add(new XlRow(sheet));

        var r = sheet.Rows[row];
        while (r.Cells.Count <= col)
            r.Cells.Add(new XlCell(r));

        var style = JsonToStyle(styleJson);
        if (style is not null)
            r.Cells[col].Style = style;
    }

    [JSExport]
    public static void SetRowStyle(int row, string styleJson)
    {
        var sheet = ActiveSheet;
        if (sheet is null) return;

        while (sheet.Rows.Count <= row)
            sheet.Rows.Add(new XlRow(sheet));

        var style = JsonToStyle(styleJson);
        if (style is not null)
            sheet.Rows[row].Style = style;
    }

    [JSExport]
    public static void SetColStyle(int col, string styleJson)
    {
        var sheet = ActiveSheet;
        if (sheet is null) return;

        if (!sheet.Columns.TryGetValue(col, out var colObj) || colObj is null)
        {
            colObj = new XlColumn();
            sheet.Columns[col] = colObj;
        }

        var style = JsonToStyle(styleJson);
        if (style is not null)
            colObj.Style = style;
    }

    // ── Layout setters ──

    [JSExport]
    public static void SetColumnWidth(int col, double width)
    {
        var sheet = ActiveSheet;
        if (sheet is null) return;

        if (!sheet.Columns.TryGetValue(col, out var colObj) || colObj is null)
        {
            colObj = new XlColumn();
            sheet.Columns[col] = colObj;
        }
        colObj.Width = width;
    }

    [JSExport]
    public static void SetRowHeight(int row, double height)
    {
        var sheet = ActiveSheet;
        if (sheet is null) return;

        while (sheet.Rows.Count <= row)
            sheet.Rows.Add(new XlRow(sheet));

        sheet.Rows[row].Height = height;
    }

    // ── Merge cells ──

    [JSExport]
    public static void MergeCells(string rangeRef)
    {
        ActiveSheet?.Range(rangeRef).Merge();
    }

    [JSExport]
    public static void UnmergeCells(string rangeRef)
    {
        ActiveSheet?.Range(rangeRef).Unmerge();
    }

    // ── Sheet management ──

    [JSExport]
    public static int AddSheet(string name)
    {
        if (_workbook is null) return -1;
        var sheet = _workbook.NewWorksheet(name);
        return _workbook.Worksheets.Count - 1;
    }

    [JSExport]
    public static bool DeleteSheet(int index)
    {
        if (_workbook is null || index < 0 || index >= _workbook.Worksheets.Count)
            return false;
        if (_workbook.Worksheets.Count <= 1)
            return false;

        _workbook.Worksheets.RemoveAt(index);
        if (_activeSheetIndex >= _workbook.Worksheets.Count)
            _activeSheetIndex = _workbook.Worksheets.Count - 1;
        return true;
    }

    [JSExport]
    public static bool RenameSheet(int index, string newName)
    {
        if (_workbook is null || index < 0 || index >= _workbook.Worksheets.Count)
            return false;

        // Check for duplicate name
        if (_workbook.Worksheets.Any(s => string.Equals(s.Name, newName, StringComparison.OrdinalIgnoreCase)))
            return false;

        _workbook.Worksheets[index].Name = newName;
        return true;
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

    private static XlStyle? JsonToStyle(string json)
    {
        if (string.IsNullOrWhiteSpace(json) || json == "{}")
            return null;

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var style = new XlStyle();

            if (root.TryGetProperty("fontName", out var el)) style.FontName = el.GetString();
            if (root.TryGetProperty("fontSize", out el) && el.TryGetDouble(out var d)) style.FontSize = d;
            if (root.TryGetProperty("fontColor", out el)) style.FontColor = CssColorToArgb(el.GetString());
            if (root.TryGetProperty("bold", out el)) style.Bold = el.GetBoolean();
            if (root.TryGetProperty("italic", out el)) style.Italic = el.GetBoolean();
            if (root.TryGetProperty("fillColor", out el)) style.FillColor = CssColorToArgb(el.GetString());

            XlAlignment? align = null;
            if (root.TryGetProperty("hAlign", out el) || root.TryGetProperty("vAlign", out el) || root.TryGetProperty("wrapText", out el))
            {
                align = new XlAlignment();
                if (root.TryGetProperty("hAlign", out el)) align.Horizontal = ParseHAlign(el.GetString());
                if (root.TryGetProperty("vAlign", out el)) align.Vertical = ParseVAlign(el.GetString());
                if (root.TryGetProperty("wrapText", out el)) align.WrapText = el.GetBoolean();
                style.Alignment = align;
            }

            // Check for border properties
            if (root.TryGetProperty("borderLeft", out el) || root.TryGetProperty("borderRight", out el) ||
                root.TryGetProperty("borderTop", out el) || root.TryGetProperty("borderBottom", out el))
            {
                var border = new XlBorder();
                if (root.TryGetProperty("borderLeft", out el)) border.Left = ParseBorderStyle(el.GetString());
                if (root.TryGetProperty("borderLeftColor", out el)) border.LeftColor = CssColorToArgb(el.GetString());
                if (root.TryGetProperty("borderRight", out el)) border.Right = ParseBorderStyle(el.GetString());
                if (root.TryGetProperty("borderRightColor", out el)) border.RightColor = CssColorToArgb(el.GetString());
                if (root.TryGetProperty("borderTop", out el)) border.Top = ParseBorderStyle(el.GetString());
                if (root.TryGetProperty("borderTopColor", out el)) border.TopColor = CssColorToArgb(el.GetString());
                if (root.TryGetProperty("borderBottom", out el)) border.Bottom = ParseBorderStyle(el.GetString());
                if (root.TryGetProperty("borderBottomColor", out el)) border.BottomColor = CssColorToArgb(el.GetString());
                style.Border = border;
            }

            return style;
        }
        catch
        {
            return null;
        }
    }

    private static string? CssColorToArgb(string? cssColor)
    {
        if (string.IsNullOrWhiteSpace(cssColor)) return null;
        // "#RRGGBB" → "FFRRGGBB"
        var hex = cssColor.TrimStart('#');
        if (hex.Length == 6) return "FF" + hex;
        if (hex.Length == 8) return hex; // Already ARGB
        return "FF" + hex.PadLeft(6, '0');
    }

    private static XlHorizontalAlignment ParseHAlign(string? s) => s switch
    {
        "left" => XlHorizontalAlignment.Left,
        "center" => XlHorizontalAlignment.Center,
        "right" => XlHorizontalAlignment.Right,
        "justify" => XlHorizontalAlignment.Justify,
        _ => XlHorizontalAlignment.General
    };

    private static XlVerticalAlignment ParseVAlign(string? s) => s switch
    {
        "top" => XlVerticalAlignment.Top,
        "middle" => XlVerticalAlignment.Center,
        "bottom" => XlVerticalAlignment.Bottom,
        "justify" => XlVerticalAlignment.Justify,
        _ => XlVerticalAlignment.Bottom
    };

    private static XlBorderStyle ParseBorderStyle(string? s) => s switch
    {
        "thin" => XlBorderStyle.Thin,
        "medium" => XlBorderStyle.Medium,
        "dashed" => XlBorderStyle.Dashed,
        "dotted" => XlBorderStyle.Dotted,
        "thick" => XlBorderStyle.Thick,
        "double" => XlBorderStyle.Double,
        "hair" => XlBorderStyle.Hair,
        _ => XlBorderStyle.None
    };
}
