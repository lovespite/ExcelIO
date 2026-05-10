namespace ExcelIO;

public class XlLoadOptions
{
    public bool LoadStyles { get; set; } = true;
    public bool LoadImages { get; set; } = true;

    public static XlLoadOptions Default => new();
}
