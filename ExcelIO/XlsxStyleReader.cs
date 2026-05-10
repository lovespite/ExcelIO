using System.IO.Compression;
using System.Xml.Linq;

namespace ExcelIO;

internal static class XlsxStyleReader
{
    private const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static List<XlStyle> ReadStyles(ZipArchive archive)
    {
        var styles = new List<XlStyle>();
        var entry = archive.GetEntry("xl/styles.xml");
        if (entry == null)
        {
            styles.Add(new XlStyle());
            return styles;
        }

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        var root = doc.Root;
        if (root == null)
        {
            styles.Add(new XlStyle());
            return styles;
        }

        var ns = XNamespace.Get(NsMain);

        // 1. Read Fonts
        var fonts = new List<XlStyle>();
        var fontsElement = root.Element(ns + "fonts");
        if (fontsElement != null)
        {
            foreach (var fontElement in fontsElement.Elements(ns + "font"))
            {
                var font = new XlStyle();
                var nameElement = fontElement.Element(ns + "name");
                if (nameElement != null) font.FontName = nameElement.Attribute("val")?.Value;

                var szElement = fontElement.Element(ns + "sz");
                if (szElement != null && double.TryParse(szElement.Attribute("val")?.Value, out var sz)) font.FontSize = sz;

                var colorElement = fontElement.Element(ns + "color");
                if (colorElement != null) font.FontColor = colorElement.Attribute("rgb")?.Value;

                if (fontElement.Element(ns + "b") != null) font.Bold = true;
                if (fontElement.Element(ns + "i") != null) font.Italic = true;

                fonts.Add(font);
            }
        }
        if (fonts.Count == 0) fonts.Add(new XlStyle());

        // 2. Read Fills
        var fills = new List<string?>();
        var fillsElement = root.Element(ns + "fills");
        if (fillsElement != null)
        {
            foreach (var fillElement in fillsElement.Elements(ns + "fill"))
            {
                var patternFill = fillElement.Element(ns + "patternFill");
                if (patternFill != null)
                {
                    var patternType = patternFill.Attribute("patternType")?.Value;
                    if (patternType != "none" && patternType != "gray125")
                    {
                        var fgColor = patternFill.Element(ns + "fgColor");
                        fills.Add(fgColor?.Attribute("rgb")?.Value);
                    }
                    else
                    {
                        fills.Add(null);
                    }
                }
                else
                {
                    fills.Add(null);
                }
            }
        }
        if (fills.Count == 0) fills.Add(null);

        // 3. Read Borders
        var borders = new List<XlBorder>();
        var bordersElement = root.Element(ns + "borders");
        if (bordersElement != null)
        {
            foreach (var borderElement in bordersElement.Elements(ns + "border"))
            {
                var border = new XlBorder();
                ReadBorderSide(borderElement, ns + "left", ref border, (b, s, c) => { b.Left = s; b.LeftColor = c; });
                ReadBorderSide(borderElement, ns + "right", ref border, (b, s, c) => { b.Right = s; b.RightColor = c; });
                ReadBorderSide(borderElement, ns + "top", ref border, (b, s, c) => { b.Top = s; b.TopColor = c; });
                ReadBorderSide(borderElement, ns + "bottom", ref border, (b, s, c) => { b.Bottom = s; b.BottomColor = c; });
                borders.Add(border);
            }
        }
        if (borders.Count == 0) borders.Add(new XlBorder());

        // 4. Read Cell Xfs
        var cellXfsElement = root.Element(ns + "cellXfs");
        if (cellXfsElement != null)
        {
            foreach (var xfElement in cellXfsElement.Elements(ns + "xf"))
            {
                var style = new XlStyle();

                var fontIdStr = xfElement.Attribute("fontId")?.Value;
                if (int.TryParse(fontIdStr, out var fontId) && fontId >= 0 && fontId < fonts.Count && xfElement.Attribute("applyFont")?.Value == "1")
                {
                    var font = fonts[fontId];
                    style.FontName = font.FontName;
                    style.FontSize = font.FontSize;
                    style.FontColor = font.FontColor;
                    style.Bold = font.Bold;
                    style.Italic = font.Italic;
                }

                var fillIdStr = xfElement.Attribute("fillId")?.Value;
                if (int.TryParse(fillIdStr, out var fillId) && fillId >= 0 && fillId < fills.Count && xfElement.Attribute("applyFill")?.Value == "1")
                {
                    style.FillColor = fills[fillId];
                }

                var borderIdStr = xfElement.Attribute("borderId")?.Value;
                if (int.TryParse(borderIdStr, out var borderId) && borderId >= 0 && borderId < borders.Count && xfElement.Attribute("applyBorder")?.Value == "1")
                {
                    style.Border = borders[borderId];
                }

                if (xfElement.Attribute("applyAlignment")?.Value == "1")
                {
                    var alignElement = xfElement.Element(ns + "alignment");
                    if (alignElement != null)
                    {
                        var alignment = new XlAlignment();
                        
                        var horiz = alignElement.Attribute("horizontal")?.Value;
                        if (Enum.TryParse<XlHorizontalAlignment>(horiz, true, out var h)) alignment.Horizontal = h;

                        var vert = alignElement.Attribute("vertical")?.Value;
                        if (Enum.TryParse<XlVerticalAlignment>(vert, true, out var v)) alignment.Vertical = v;

                        if (alignElement.Attribute("wrapText")?.Value == "1") alignment.WrapText = true;

                        style.Alignment = alignment;
                    }
                }

                styles.Add(style);
            }
        }
        else
        {
            styles.Add(new XlStyle());
        }

        return styles;
    }

    private delegate void SetBorder(XlBorder border, XlBorderStyle style, string? color);

    private static void ReadBorderSide(XElement borderElement, XName sideName, ref XlBorder border, SetBorder setter)
    {
        var sideElement = borderElement.Element(sideName);
        if (sideElement != null)
        {
            var styleStr = sideElement.Attribute("style")?.Value;
            if (!string.IsNullOrEmpty(styleStr) && Enum.TryParse<XlBorderStyle>(styleStr, true, out var style))
            {
                var colorElement = sideElement.Element(sideName.Namespace + "color");
                var color = colorElement?.Attribute("rgb")?.Value;
                setter(border, style, color);
            }
        }
    }
}
