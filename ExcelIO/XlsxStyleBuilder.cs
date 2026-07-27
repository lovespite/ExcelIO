using System.Text;

namespace ExcelIO;

internal class XlsxStyleBuilder
{
    private readonly List<XlStyle> _styles = [];
    private readonly List<XlFont> _fonts = [];
    private readonly List<XlFill> _fills = [];
    private readonly List<XlBorder> _borders = [];

    // Dictionary for O(1) lookup
    private readonly Dictionary<XlStyle, int> _styleIndexMap = new();
    private readonly Dictionary<XlFont, int> _fontIndexMap = new();
    private readonly Dictionary<XlFill, int> _fillIndexMap = new();
    private readonly Dictionary<XlBorder, int> _borderIndexMap = new();

    public XlsxStyleBuilder()
    {
        // Default styles (index 0)
        var defaultFont = new XlFont();
        _fonts.Add(defaultFont);
        _fontIndexMap[defaultFont] = 0;

        var defaultFill0 = new XlFill { PatternType = "none" };
        _fills.Add(defaultFill0);
        _fillIndexMap[defaultFill0] = 0;

        var defaultFill1 = new XlFill { PatternType = "gray125" };
        _fills.Add(defaultFill1);
        _fillIndexMap[defaultFill1] = 1;

        var defaultBorder = new XlBorder();
        _borders.Add(defaultBorder);
        _borderIndexMap[defaultBorder] = 0;

        var defaultStyle = new XlStyle();
        _styles.Add(defaultStyle);
        _styleIndexMap[defaultStyle] = 0;
    }

    public int GetStyleIndex(XlStyle? style)
    {
        if (style == null) return 0;

        // Check if style already exists
        if (_styleIndexMap.TryGetValue(style, out var existingIndex))
            return existingIndex;

        // New style: get or create component IDs
        GetOrAddFont(style);
        GetOrAddFill(style);
        GetOrAddBorder(style);

        // Register new style
        int index = _styles.Count;
        _styles.Add(style);
        _styleIndexMap[style] = index;
        return index;
    }

    private int GetOrAddFont(XlStyle style)
    {
        var font = new XlFont
        {
            Name = style.FontName ?? "Calibri",
            Size = style.FontSize ?? 11,
            Color = style.FontColor,
            Bold = style.Bold,
            Italic = style.Italic
        };

        if (_fontIndexMap.TryGetValue(font, out var index))
            return index;

        index = _fonts.Count;
        _fonts.Add(font);
        _fontIndexMap[font] = index;
        return index;
    }

    private int GetOrAddFill(XlStyle style)
    {
        var fill = new XlFill
        {
            PatternType = style.FillColor != null ? "solid" : "none",
            ForegroundColor = style.FillColor
        };

        if (_fillIndexMap.TryGetValue(fill, out var index))
            return index;

        index = _fills.Count;
        _fills.Add(fill);
        _fillIndexMap[fill] = index;
        return index;
    }

    private int GetOrAddBorder(XlStyle style)
    {
        var border = style.Border ?? new XlBorder();

        if (_borderIndexMap.TryGetValue(border, out var index))
            return index;

        index = _borders.Count;
        _borders.Add(border);
        _borderIndexMap[border] = index;
        return index;
    }

    // Helper methods for GenerateStylesXml to lookup component IDs
    private int GetFontIdForStyle(XlStyle style)
    {
        var font = new XlFont
        {
            Name = style.FontName ?? "Calibri",
            Size = style.FontSize ?? 11,
            Color = style.FontColor,
            Bold = style.Bold,
            Italic = style.Italic
        };
        return _fontIndexMap.TryGetValue(font, out var id) ? id : 0;
    }

    private int GetFillIdForStyle(XlStyle style)
    {
        var fill = new XlFill
        {
            PatternType = style.FillColor != null ? "solid" : "none",
            ForegroundColor = style.FillColor
        };
        return _fillIndexMap.TryGetValue(fill, out var id) ? id : 0;
    }

    private int GetBorderIdForStyle(XlStyle style)
    {
        var border = style.Border ?? new XlBorder();
        return _borderIndexMap.TryGetValue(border, out var id) ? id : 0;
    }

    public string GenerateStylesXml()
    {
        var sb = new StringBuilder(capacity: (_fonts.Count + _fills.Count + _borders.Count + _styles.Count) * 150 + 2048);
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
            int fontId = GetFontIdForStyle(style);
            int fillId = GetFillIdForStyle(style);
            int borderId = GetBorderIdForStyle(style);

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
