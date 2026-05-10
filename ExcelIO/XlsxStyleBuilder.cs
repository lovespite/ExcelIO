using System.Text;

namespace ExcelIO;

internal class XlsxStyleBuilder
{
    private readonly List<XlStyle> _styles = [];
    private readonly List<XlFont> _fonts = [];
    private readonly List<XlFill> _fills = [];
    private readonly List<XlBorder> _borders = [];

    public XlsxStyleBuilder()
    {
        // Default styles (index 0)
        _fonts.Add(new XlFont()); // Default font
        _fills.Add(new XlFill { PatternType = "none" }); // Default fill 0: none
        _fills.Add(new XlFill { PatternType = "gray125" }); // Default fill 1: gray125
        _borders.Add(new XlBorder()); // Default border
        _styles.Add(new XlStyle()); // Default XF
    }

    public int GetStyleIndex(XlStyle? style)
    {
        if (style == null) return 0;

        // Simplify style for matching (only properties that affect styles.xml)
        var font = new XlFont
        {
            Name = style.FontName ?? "Calibri",
            Size = style.FontSize ?? 11,
            Color = style.FontColor,
            Bold = style.Bold,
            Italic = style.Italic
        };

        var fill = new XlFill
        {
            PatternType = style.FillColor != null ? "solid" : "none",
            ForegroundColor = style.FillColor
        };

        var border = style.Border ?? new XlBorder();

        int fontId = GetOrAdd(_fonts, font);
        int fillId = GetOrAdd(_fills, fill);
        int borderId = GetOrAdd(_borders, border);

        // Check if this combination already exists
        for (int i = 0; i < _styles.Count; i++)
        {
            var s = _styles[i];
            if (fontId == GetFontId(s) && fillId == GetFillId(s) && borderId == GetBorderId(s) &&
                Equals(s.Alignment, style.Alignment))
            {
                return i;
            }
        }

        _styles.Add(style);
        return _styles.Count - 1;
    }

    private int GetFontId(XlStyle style)
    {
        var font = new XlFont
        {
            Name = style.FontName ?? "Calibri",
            Size = style.FontSize ?? 11,
            Color = style.FontColor,
            Bold = style.Bold,
            Italic = style.Italic
        };
        return _fonts.IndexOf(font);
    }

    private int GetFillId(XlStyle style)
    {
        var fill = new XlFill
        {
            PatternType = style.FillColor != null ? "solid" : "none",
            ForegroundColor = style.FillColor
        };
        return _fills.IndexOf(fill);
    }

    private int GetBorderId(XlStyle style)
    {
        return _borders.IndexOf(style.Border ?? new XlBorder());
    }

    private static int GetOrAdd<T>(List<T> list, T item) where T : IEquatable<T>
    {
        int index = list.IndexOf(item);
        if (index == -1)
        {
            list.Add(item);
            return list.Count - 1;
        }
        return index;
    }

    public string GenerateStylesXml()
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

        // Fonts
        sb.Append($"<fonts count=\"{_fonts.Count}\">");
        foreach (var font in _fonts)
        {
            sb.Append("<font>");
            if (font.Bold) sb.Append("<b/>");
            if (font.Italic) sb.Append("<i/>");
            sb.Append($"<sz val=\"{font.Size}\"/>");
            if (!string.IsNullOrEmpty(font.Color)) sb.Append($"<color rgb=\"{font.Color}\"/>");
            sb.Append($"<name val=\"{font.Name}\"/>");
            sb.Append("<family val=\"2\"/><scheme val=\"minor\"/>");
            sb.Append("</font>");
        }
        sb.Append("</fonts>");

        // Fills
        sb.Append($"<fills count=\"{_fills.Count}\">");
        foreach (var fill in _fills)
        {
            sb.Append($"<fill><patternFill patternType=\"{fill.PatternType}\">");
            if (!string.IsNullOrEmpty(fill.ForegroundColor))
            {
                sb.Append($"<fgColor rgb=\"{fill.ForegroundColor}\"/>");
            }
            sb.Append("</patternFill></fill>");
        }
        sb.Append("</fills>");

        // Borders
        sb.Append($"<borders count=\"{_borders.Count}\">");
        foreach (var border in _borders)
        {
            sb.Append("<border>");
            AppendBorderSide(sb, "left", border.Left, border.LeftColor);
            AppendBorderSide(sb, "right", border.Right, border.RightColor);
            AppendBorderSide(sb, "top", border.Top, border.TopColor);
            AppendBorderSide(sb, "bottom", border.Bottom, border.BottomColor);
            sb.Append("<diagonal/>");
            sb.Append("</border>");
        }
        sb.Append("</borders>");

        // Cell Style Xfs (Default)
        sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

        // Cell Xfs
        sb.Append($"<cellXfs count=\"{_styles.Count}\">");
        for (int i = 0; i < _styles.Count; i++)
        {
            var style = _styles[i];
            int fontId = GetFontId(style);
            int fillId = GetFillId(style);
            int borderId = GetBorderId(style);

            sb.Append($"<xf numFmtId=\"0\" fontId=\"{fontId}\" fillId=\"{fillId}\" borderId=\"{borderId}\" xfId=\"0\"");
            if (fontId > 0) sb.Append(" applyFont=\"1\"");
            if (fillId > 0) sb.Append(" applyFill=\"1\"");
            if (borderId > 0) sb.Append(" applyBorder=\"1\"");
            if (style.Alignment != null) sb.Append(" applyAlignment=\"1\"");

            if (style.Alignment != null)
            {
                sb.Append(">");
                sb.Append($"<alignment horizontal=\"{style.Alignment.Horizontal.ToString().ToLowerInvariant()}\" vertical=\"{style.Alignment.Vertical.ToString().ToLowerInvariant()}\"");
                if (style.Alignment.WrapText) sb.Append(" wrapText=\"1\"");
                sb.Append("/>");
                sb.Append("</xf>");
            }
            else
            {
                sb.Append("/>");
            }
        }
        sb.Append("</cellXfs>");

        sb.Append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
        sb.Append("<dxfs count=\"0\"/>");
        sb.Append("<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>");
        sb.Append("</styleSheet>");

        return sb.ToString();
    }

    private static void AppendBorderSide(StringBuilder sb, string side, XlBorderStyle style, string? color)
    {
        if (style == XlBorderStyle.None)
        {
            sb.Append($"<{side}/>");
        }
        else
        {
            sb.Append($"<{side} style=\"{style.ToString().ToLowerInvariant()}\">");
            if (!string.IsNullOrEmpty(color)) sb.Append($"<color rgb=\"{color}\"/>");
            sb.Append($"</{side}>");
        }
    }

    private class XlFont : IEquatable<XlFont>
    {
        public string Name { get; set; } = "Calibri";
        public double Size { get; set; } = 11;
        public string? Color { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }

        public bool Equals(XlFont? other)
        {
            if (other is null) return false;
            return Name == other.Name && Size == other.Size && Color == other.Color && Bold == other.Bold && Italic == other.Italic;
        }

        public override bool Equals(object? obj) => Equals(obj as XlFont);
        public override int GetHashCode() => HashCode.Combine(Name, Size, Color, Bold, Italic);
    }

    private class XlFill : IEquatable<XlFill>
    {
        public string PatternType { get; set; } = "none";
        public string? ForegroundColor { get; set; }

        public bool Equals(XlFill? other)
        {
            if (other is null) return false;
            return PatternType == other.PatternType && ForegroundColor == other.ForegroundColor;
        }

        public override bool Equals(object? obj) => Equals(obj as XlFill);
        public override int GetHashCode() => HashCode.Combine(PatternType, ForegroundColor);
    }
}
