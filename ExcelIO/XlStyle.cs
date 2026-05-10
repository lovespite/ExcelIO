namespace ExcelIO;

public class XlStyle : IEquatable<XlStyle>
{
    public string? FontName { get; set; }
    public double? FontSize { get; set; }
    public string? FontColor { get; set; } // ARGB hex string, e.g., "FFFF0000"
    public bool Bold { get; set; }
    public bool Italic { get; set; }

    public string? FillColor { get; set; } // ARGB hex string

    public XlAlignment? Alignment { get; set; }

    public XlBorder? Border { get; set; }

    public bool Equals(XlStyle? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return FontName == other.FontName &&
               FontSize == other.FontSize &&
               FontColor == other.FontColor &&
               Bold == other.Bold &&
               Italic == other.Italic &&
               FillColor == other.FillColor &&
               Equals(Alignment, other.Alignment) &&
               Equals(Border, other.Border);
    }

    public override bool Equals(object? obj) => Equals(obj as XlStyle);

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(FontName);
        hash.Add(FontSize);
        hash.Add(FontColor);
        hash.Add(Bold);
        hash.Add(Italic);
        hash.Add(FillColor);
        hash.Add(Alignment);
        hash.Add(Border);
        return hash.ToHashCode();
    }
}

public class XlAlignment : IEquatable<XlAlignment>
{
    public XlHorizontalAlignment Horizontal { get; set; } = XlHorizontalAlignment.General;
    public XlVerticalAlignment Vertical { get; set; } = XlVerticalAlignment.Bottom;
    public bool WrapText { get; set; }

    public bool Equals(XlAlignment? other)
    {
        if (other is null) return false;
        return Horizontal == other.Horizontal && Vertical == other.Vertical && WrapText == other.WrapText;
    }

    public override bool Equals(object? obj) => Equals(obj as XlAlignment);

    public override int GetHashCode() => HashCode.Combine(Horizontal, Vertical, WrapText);
}

public enum XlHorizontalAlignment
{
    General,
    Left,
    Center,
    Right,
    Justify
}

public enum XlVerticalAlignment
{
    Top,
    Center,
    Bottom,
    Justify
}

public class XlBorder : IEquatable<XlBorder>
{
    public XlBorderStyle Left { get; set; } = XlBorderStyle.None;
    public XlBorderStyle Right { get; set; } = XlBorderStyle.None;
    public XlBorderStyle Top { get; set; } = XlBorderStyle.None;
    public XlBorderStyle Bottom { get; set; } = XlBorderStyle.None;

    public string? LeftColor { get; set; }
    public string? RightColor { get; set; }
    public string? TopColor { get; set; }
    public string? BottomColor { get; set; }

    public bool Equals(XlBorder? other)
    {
        if (other is null) return false;
        return Left == other.Left && Right == other.Right && Top == other.Top && Bottom == other.Bottom &&
               LeftColor == other.LeftColor && RightColor == other.RightColor && TopColor == other.TopColor && BottomColor == other.BottomColor;
    }

    public override bool Equals(object? obj) => Equals(obj as XlBorder);

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(Left);
        hash.Add(Right);
        hash.Add(Top);
        hash.Add(Bottom);
        hash.Add(LeftColor);
        hash.Add(RightColor);
        hash.Add(TopColor);
        hash.Add(BottomColor);
        return hash.ToHashCode();
    }
}

public enum XlBorderStyle
{
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair
}
