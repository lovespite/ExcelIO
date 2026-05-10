namespace ExcelIO;

public sealed class XlWorksheetImage
{
    private static readonly HashSet<string> SupportedExtensions =
    [
        "png",
        "jpg",
        "jpeg",
        "gif",
        "bmp",
        "tif",
        "tiff"
    ];

    private XlWorksheetImage(byte[] imageBytes, string extension, int rowIndex, int columnIndex, int rowSpan, int columnSpan)
    {
        Bytes = imageBytes;
        Extension = extension;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        RowSpan = rowSpan;
        ColumnSpan = columnSpan;
    }

    public byte[] Bytes { get; }
    public string Extension { get; }
    public int RowIndex { get; }
    public int ColumnIndex { get; }
    public int RowSpan { get; }
    public int ColumnSpan { get; }

    internal static XlWorksheetImage Create(byte[] imageBytes, string imageExtension, int rowIndex, int columnIndex, int rowSpan, int columnSpan)
    {
        ArgumentNullException.ThrowIfNull(imageBytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(imageExtension);

        if (imageBytes.Length == 0)
        {
            throw new ArgumentException("Image bytes cannot be empty.", nameof(imageBytes));
        }
        if (rowIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than or equal to 0.");
        }
        if (columnIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be greater than or equal to 0.");
        }
        if (rowSpan <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rowSpan), "Row span must be greater than 0.");
        }
        if (columnSpan <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columnSpan), "Column span must be greater than 0.");
        }

        var normalizedExtension = NormalizeExtension(imageExtension);
        if (!SupportedExtensions.Contains(normalizedExtension))
        {
            throw new NotSupportedException($"Unsupported image format: .{normalizedExtension}");
        }

        return new XlWorksheetImage([.. imageBytes], normalizedExtension, rowIndex, columnIndex, rowSpan, columnSpan);
    }

    private static string NormalizeExtension(string imageExtension)
    {
        var normalized = imageExtension.Trim();
        if (normalized.StartsWith('.'))
        {
            normalized = normalized[1..];
        }
        return normalized.ToLowerInvariant();
    }
}
