namespace ExcelIO;

public class XlSheetOptions
{
    public string? TabColor { get; set; } // ARGB hex string
    public bool ShowGridLines { get; set; } = true;
    public double? DefaultRowHeight { get; set; }
}
