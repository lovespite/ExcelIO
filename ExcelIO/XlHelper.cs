using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ExcelIO;

public static class XlHelper
{
    // OpenXML 命名空间常量
    private const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static bool IsXlsxFile(string filepath)
    {
        var ext = Path.GetExtension(filepath);
        return string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(ext, ".xlsm", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsXlsFile(string filepath)
    {
        var ext = Path.GetExtension(filepath);
        return string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase);
    }

    public static Task SaveAsync(string filepath, XlWorkbook workbookData)
        => Task.Run(() => Save(filepath, workbookData));

    public static Task<XlWorkbook> LoadAsync(string filepath, XlLoadOptions? options = null)
        => Task.Run(() => Load(filepath, options));

    public static async Task<XlWorkbook> LoadAsync(Stream stream, string extension, XlLoadOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return LoadByExtension(stream, extension, options);
    }

    /// <summary>
    /// 保存 Excel (OpenXML 格式)
    /// </summary>
    public static void Save(string filepath, XlWorkbook workbookData)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filepath);
        ArgumentNullException.ThrowIfNull(workbookData);

        if (File.Exists(filepath)) File.Delete(filepath);

        var drawingSheets = new List<(int SheetIndex, int DrawingIndex, XlWorksheet Sheet)>();
        var cellImages = new List<(XlWorksheet Sheet, XlWorksheetImage Image, int vmIndex, int richValueIndex, int imageIndex)>();
        var imageExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        
        int drawingIndex = 1;
        int globalImageIndex = 1;

        for (int i = 0; i < workbookData.Worksheets.Count; i++)
        {
            var sheet = workbookData.Worksheets[i];
            var floatingImages = sheet.Images.Where(img => !img.PlaceInCell).ToList();
            var sheetCellImages = sheet.Images.Where(img => img.PlaceInCell).ToList();

            if (floatingImages.Count > 0)
            {
                drawingSheets.Add((i + 1, drawingIndex, sheet));
                drawingIndex++;
            }

            foreach (var img in sheetCellImages)
            {
                // Ensure the cell exists so it gets rendered in sheet.xml
                while (sheet.Rows.Count <= img.RowIndex) sheet.Rows.Add(new XlRow(sheet));
                var row = sheet.Rows[img.RowIndex];
                while (row.Cells.Count <= img.ColumnIndex) row.Cells.Add(new XlCell(row));

                // vmIndex is 1-based, richValueIndex and imageIndex are 0-based
                int currentCellImageCount = cellImages.Count;
                cellImages.Add((sheet, img, currentCellImageCount + 1, currentCellImageCount, currentCellImageCount));
                imageExtensions.Add(img.Extension);
            }
            
            foreach (var img in floatingImages)
            {
                imageExtensions.Add(img.Extension);
            }
        }
        var drawingSheetMap = drawingSheets.ToDictionary(x => x.SheetIndex, x => x.DrawingIndex);

        using var fileStream = new FileStream(filepath, FileMode.Create);
        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Create);

        // 1. 写入 [Content_Types].xml (定义文件类型)
        WriteEntry(archive, "[Content_Types].xml", GenerateContentTypes(workbookData.Worksheets.Count, drawingSheets.Count, cellImages.Count > 0, imageExtensions));

        // 2. 写入 _rels/.rels (定义根关系)
        WriteEntry(archive, "_rels/.rels", GenerateRootRels());

        // 3. 写入 xl/workbook.xml (定义工作簿结构)
        WriteEntry(archive, "xl/workbook.xml", GenerateWorkbookXml(workbookData.Worksheets));

        // 4. 写入 xl/_rels/workbook.xml.rels (定义工作簿与Sheet的关系)
        WriteEntry(archive, "xl/_rels/workbook.xml.rels", GenerateWorkbookRels(workbookData.Worksheets.Count, cellImages.Count > 0));

        var styleBuilder = new XlsxStyleBuilder();

        // 6. 写入具体的 Sheet 数据
        for (int i = 0; i < workbookData.Worksheets.Count; i++)
        {
            var sheet = workbookData.Worksheets[i];
            drawingSheetMap.TryGetValue(i + 1, out int sheetDrawingIndex);
            var path = "xl/worksheets/sheet" + (i + 1) + ".xml";
            
            var sheetVmMap = cellImages.Where(x => x.Sheet == sheet)
                                       .ToDictionary(x => (x.Image.RowIndex, x.Image.ColumnIndex), x => x.vmIndex);

            WriteEntry(archive, path, GenerateSheetXml(sheet, styleBuilder, sheetVmMap, sheetDrawingIndex > 0 ? "rId1" : null));
            if (sheetDrawingIndex > 0)
            {
                WriteEntry(archive, "xl/worksheets/_rels/sheet" + (i + 1) + ".xml.rels", GenerateWorksheetRels(sheetDrawingIndex));
            }
        }

        // 5. 写入 xl/styles.xml
        WriteEntry(archive, "xl/styles.xml", styleBuilder.GenerateStylesXml());

        // 7. 写入 Rich Data 文件
        if (cellImages.Count > 0)
        {
            WriteEntry(archive, "xl/metadata.xml", GenerateMetadataXml(cellImages.Count));
            WriteEntry(archive, "xl/richData/rdrichvaluestructure.xml", GenerateRichValueStructureXml());
            WriteEntry(archive, "xl/richData/rdRichValueTypes.xml", GenerateRichValueTypesXml());
            
            var rvParts = new List<RichValuePart>();
            var rvRelParts = new List<RichValueRelPart>();
            
            for (int ci = 0; ci < cellImages.Count; ci++)
            {
                var mediaFileName = "image" + globalImageIndex + "." + cellImages[ci].Image.Extension;
                WriteBinaryEntry(archive, "xl/media/" + mediaFileName, cellImages[ci].Image.Bytes);
                
                rvParts.Add(new RichValuePart(cellImages[ci].richValueIndex, cellImages[ci].imageIndex));
                rvRelParts.Add(new RichValueRelPart("rId" + (ci + 1), mediaFileName));
                
                globalImageIndex++;
            }
            
            WriteEntry(archive, "xl/richData/rdrichvalue.xml", GenerateRichValueXml(rvParts));
            WriteEntry(archive, "xl/richData/richValueRel.xml", GenerateRichValueRelXml(rvRelParts));
            WriteEntry(archive, "xl/richData/_rels/richValueRel.xml.rels", GenerateRichValueRelRels(rvRelParts));
        }

        foreach (var (_, currentDrawingIndex, sheet) in drawingSheets)
        {
            var floatingImages = sheet.Images.Where(img => !img.PlaceInCell).ToList();
            var drawingImageParts = new List<DrawingImagePart>(floatingImages.Count);
            for (int imageIndex = 0; imageIndex < floatingImages.Count; imageIndex++)
            {
                var image = floatingImages[imageIndex];
                var mediaFileName = "image" + globalImageIndex + "." + image.Extension;
                WriteBinaryEntry(archive, "xl/media/" + mediaFileName, image.Bytes);
                drawingImageParts.Add(new DrawingImagePart("rId" + (imageIndex + 1), mediaFileName, image, imageIndex + 1));
                globalImageIndex++;
            }

            WriteEntry(archive, "xl/drawings/drawing" + currentDrawingIndex + ".xml", GenerateDrawingXml(drawingImageParts));
            WriteEntry(archive, "xl/drawings/_rels/drawing" + currentDrawingIndex + ".xml.rels", GenerateDrawingRels(drawingImageParts));
        }
    }

    public static XlWorkbook Load(string filepath, XlLoadOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filepath);
        using var fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return LoadByExtension(fs, filepath, options);
    }

    public static XlWorkbook Load(Stream stream, XlLoadOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return LoadBySignature(stream, options);
    }

    public static XlWorkbook Load(Stream stream, ExcelFormat format, XlLoadOptions? options = null)
    {
        return format switch
        {
            ExcelFormat.OpenXmlZip => LoadOpenXml(stream, options),
            ExcelFormat.Cfb => LoadByExtension(stream, ".xls", options),
            _ => LoadBySignature(stream, options),
        };
    }

    private static XlWorkbook LoadByExtension(Stream stream, string extension, XlLoadOptions? options = null)
    {
        if (IsXlsFile(extension))
        {
            var xlsBytes = XlsCompoundReader.ReadWorkbookBytes(stream);
            var workbookStream = XlsCompoundReader.ReadWorkbookStream(xlsBytes);
            return XlsBiff8Reader.Load(workbookStream);
        }
        if (!IsXlsxFile(extension))
        {
            throw new NotSupportedException($"Unsupported Excel format: {Path.GetExtension(extension)}");
        }

        return LoadOpenXml(stream, options);
    }

    private static XlWorkbook LoadBySignature(Stream stream, XlLoadOptions? options = null)
    {
        byte[] header = new byte[8];
        int bytesRead = stream.Read(header, 0, header.Length);
        if (bytesRead < 4)
        {
            throw new NotSupportedException("The provided stream does not appear to be a valid Excel file format.");
        }

        var format = DetectFormatFromHeader(header.AsSpan(0, bytesRead));
        if (format == ExcelFormat.Unknown)
        {
            throw new NotSupportedException("The provided stream does not appear to be a valid Excel file format.");
        }

        try
        {
            stream.Seek(0, SeekOrigin.Begin);
            return Load(stream, format, options);
        }
        catch (NotSupportedException)
        {
            using var ms = new MemoryStream();
            ms.Write(header, 0, bytesRead);
            stream.CopyTo(ms);
            ms.Seek(0, SeekOrigin.Begin);
            return Load(ms, format, options);
        }
    }

    private static XlWorkbook LoadOpenXml(Stream stream, XlLoadOptions? options = null)
    {
        options ??= XlLoadOptions.Default;
        var result = new XlWorkbook();

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        // 1. 读取 SharedStrings
        var sharedStrings = new List<string>();
        var sharedStringEntry = archive.GetEntry("xl/sharedStrings.xml");
        if (sharedStringEntry != null)
        {
            using var entryStream = sharedStringEntry.Open();
            var doc = XDocument.Load(entryStream);
            foreach (var si in doc.Descendants(XName.Get("si", NsMain)))
            {
                var val = si.DescendantNodes().OfType<XText>().Select(x => x.Value).Aggregate(new StringBuilder(), (sb, v) => sb.Append(v)).ToString();
                sharedStrings.Add(val);
            }
        }

        // 2. 读取 Styles
        var styles = options.LoadStyles ? XlsxStyleReader.ReadStyles(archive) : [new XlStyle()];

        // 3. 读取 Workbook
        var sheetMapping = new List<(string Name, string RelId)>();
        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        if (workbookEntry != null)
        {
            using var entryStream = workbookEntry.Open();
            var doc = XDocument.Load(entryStream);
            var sheets = doc.Descendants(XName.Get("sheet", NsMain));
            foreach (var sheet in sheets)
            {
                var name = sheet.Attribute("name")?.Value ?? "Unknown";
                var rid = sheet.Attribute(XName.Get("id", NsRel))?.Value;
                if (rid != null)
                {
                    sheetMapping.Add((name, rid));
                }
            }
        }

        // 4. 解析 Workbook 关系文件
        var relMapping = new Dictionary<string, string>();
        var relEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        if (relEntry != null)
        {
            using var entryStream = relEntry.Open();
            var doc = XDocument.Load(entryStream);
            foreach (var rel in doc.Descendants(XName.Get("Relationship", NsPkgRel)))
            {
                var id = rel.Attribute("Id")?.Value;
                var target = rel.Attribute("Target")?.Value;
                if (id != null && target != null)
                {
                    relMapping[id] = target;
                }
            }
        }

        // 5. 遍历并读取每个 Sheet 的数据
        foreach (var (sheetName, rid) in sheetMapping)
        {
            if (!relMapping.TryGetValue(rid, out var targetPath)) continue;

            string entryPath = targetPath.StartsWith("/") ? targetPath.TrimStart('/') : $"xl/{targetPath}";
            if (!entryPath.StartsWith("xl/")) entryPath = "xl/" + entryPath;

            var sheetEntry = archive.GetEntry(entryPath);
            if (sheetEntry == null) sheetEntry = archive.GetEntry(targetPath);
            if (sheetEntry == null) continue;

            var ws = new XlWorksheet(result) { Name = sheetName };
            result.Worksheets.Add(ws);

            string? drawingRid = null;

            using (var entryStream = sheetEntry.Open())
            using (var reader = XmlReader.Create(entryStream))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (reader.LocalName == "tabColor")
                        {
                            ws.Options.TabColor = reader.GetAttribute("rgb");
                        }
                        else if (reader.LocalName == "sheetView")
                        {
                            var showGrid = reader.GetAttribute("showGridLines");
                            if (showGrid == "0") ws.Options.ShowGridLines = false;
                        }
                        else if (reader.LocalName == "sheetFormatPr")
                        {
                            var defHeight = reader.GetAttribute("defaultRowHeight");
                            if (double.TryParse(defHeight, out var h)) ws.Options.DefaultRowHeight = h;
                        }
                        else if (reader.LocalName == "col")
                        {
                            var min = int.Parse(reader.GetAttribute("min") ?? "1") - 1;
                            var max = int.Parse(reader.GetAttribute("max") ?? "1") - 1;
                            var width = reader.GetAttribute("width");
                            var hidden = reader.GetAttribute("hidden") == "1";
                            var styleIdxStr = reader.GetAttribute("style");
                            
                            for (int c = min; c <= max; c++)
                            {
                                var col = new XlColumn { Hidden = hidden };
                                if (double.TryParse(width, out var w)) col.Width = w;
                                if (options.LoadStyles && int.TryParse(styleIdxStr, out var sIdx) && sIdx < styles.Count) col.Style = styles[sIdx];
                                ws.Columns[c] = col;
                            }
                        }
                        else if (reader.LocalName == "row")
                        {
                            var rIndexAttr = reader.GetAttribute("r");
                            if (int.TryParse(rIndexAttr, out int rIndex))
                            {
                                rIndex -= 1;
                                while (ws.Rows.Count < rIndex) ws.Rows.Add(new XlRow(ws));
                                
                                var row = new XlRow(ws);
                                ws.Rows.Add(row);

                                var ht = reader.GetAttribute("ht");
                                if (double.TryParse(ht, out var h)) row.Height = h;
                                row.Hidden = reader.GetAttribute("hidden") == "1";
                                
                                if (options.LoadStyles)
                                {
                                    var sIdxStr = reader.GetAttribute("s");
                                    if (int.TryParse(sIdxStr, out var sIdx) && sIdx < styles.Count) row.Style = styles[sIdx];
                                }
                            }
                        }
                        else if (reader.LocalName == "c")
                        {
                            var currentRow = ws.Rows.LastOrDefault();
                            if (currentRow == null) continue;

                            var type = reader.GetAttribute("t");
                            var sIdxStr = reader.GetAttribute("s");
                            XlStyle? cellStyle = null;
                            if (options.LoadStyles && int.TryParse(sIdxStr, out var sIdx) && sIdx < styles.Count) cellStyle = styles[sIdx];

                            string cellValue = string.Empty;
                            using (var subReader = reader.ReadSubtree())
                            {
                                while (subReader.Read())
                                {
                                    if (subReader.NodeType == XmlNodeType.Element)
                                    {
                                        if (subReader.LocalName == "v")
                                        {
                                            var raw = subReader.ReadElementContentAsString();
                                            if (type == "s" && int.TryParse(raw, out int idx) && idx < sharedStrings.Count)
                                                cellValue = sharedStrings[idx];
                                            else if (type == "b")
                                                cellValue = raw == "1" ? "TRUE" : "FALSE";
                                            else
                                                cellValue = raw;
                                        }
                                        else if (subReader.LocalName == "t")
                                        {
                                            cellValue = subReader.ReadElementContentAsString();
                                        }
                                    }
                                }
                            }
                            currentRow.Cells.Add(new XlCell(currentRow) { Value = cellValue, Style = cellStyle });
                        }
                        else if (reader.LocalName == "drawing")
                        {
                            if (options.LoadImages)
                            {
                                drawingRid = reader.GetAttribute("id", NsRel);
                                if (drawingRid == null) drawingRid = reader.GetAttribute("r:id");
                            }
                        }
                        else if (reader.LocalName == "mergeCell")
                        {
                            var r = reader.GetAttribute("ref");
                            if (r != null) ws.MergedCells.Add(r);
                        }
                    }
                }
            }

            // 6. 读取 Images
            if (options.LoadImages && drawingRid != null)
            {
                ReadSheetImages(archive, entryPath, drawingRid, ws);
            }
        }

        return result;
    }

    private static void ReadSheetImages(ZipArchive archive, string sheetPath, string drawingRid, XlWorksheet ws)
    {
        // 1. Find drawing path from sheet rels
        var sheetDir = Path.GetDirectoryName(sheetPath)?.Replace("\\", "/") ?? "";
        if (string.IsNullOrEmpty(sheetDir)) sheetDir = ".";
        var relPath = $"{sheetDir}/_rels/{Path.GetFileName(sheetPath)}.rels";
        var relEntry = archive.GetEntry(relPath);
        if (relEntry == null) return;

        string? drawingPath = null;
        using (var stream = relEntry.Open())
        {
            var doc = XDocument.Load(stream);
            var rel = doc.Descendants(XName.Get("Relationship", NsPkgRel))
                         .FirstOrDefault(r => r.Attribute("Id")?.Value == drawingRid);
            drawingPath = rel?.Attribute("Target")?.Value;
        }

        if (drawingPath == null) return;
        
        // Resolve Target relative to sheetDir
        if (drawingPath.StartsWith("/")) 
            drawingPath = drawingPath.TrimStart('/');
        else if (!drawingPath.StartsWith("xl/"))
        {
            var baseUri = new Uri("file:///xl/worksheets/sheet.xml");
            if (sheetPath.StartsWith("xl/")) baseUri = new Uri("file:///" + sheetPath);
            var targetUri = new Uri(baseUri, drawingPath);
            drawingPath = targetUri.AbsolutePath.TrimStart('/');
        }

        var drawingEntry = archive.GetEntry(drawingPath);
        if (drawingEntry == null) return;

        // 2. Find drawing relationships
        var drawingDir = Path.GetDirectoryName(drawingPath)?.Replace("\\", "/") ?? "";
        if (string.IsNullOrEmpty(drawingDir)) drawingDir = ".";
        var drawingRelPath = $"{drawingDir}/_rels/{Path.GetFileName(drawingPath)}.rels";
        var drawingRelEntry = archive.GetEntry(drawingRelPath);
        var imgRelMap = new Dictionary<string, string>();
        if (drawingRelEntry != null)
        {
            using var stream = drawingRelEntry.Open();
            var doc = XDocument.Load(stream);
            foreach (var rel in doc.Descendants(XName.Get("Relationship", NsPkgRel)))
            {
                var id = rel.Attribute("Id")?.Value;
                var target = rel.Attribute("Target")?.Value;
                if (id != null && target != null) imgRelMap[id] = target;
            }
        }

        // 3. Parse Drawing XML
        using (var stream = drawingEntry.Open())
        {
            var doc = XDocument.Load(stream);
            var nsDr = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            var nsA = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
            var nsR = XNamespace.Get(NsRel);

            foreach (var anchor in doc.Descendants(nsDr + "twoCellAnchor"))
            {
                var from = anchor.Element(nsDr + "from");
                var to = anchor.Element(nsDr + "to");
                if (from == null || to == null) continue;

                int fromCol = int.Parse(from.Element(nsDr + "col")?.Value ?? "0");
                int fromRow = int.Parse(from.Element(nsDr + "row")?.Value ?? "0");
                int toCol = int.Parse(to.Element(nsDr + "col")?.Value ?? "0");
                int toRow = int.Parse(to.Element(nsDr + "row")?.Value ?? "0");

                var blip = anchor.Descendants(nsA + "blip").FirstOrDefault();
                var embedId = blip?.Attribute(nsR + "embed")?.Value;
                if (embedId != null && imgRelMap.TryGetValue(embedId, out var imgPath))
                {
                    string fullImgPath;
                    if (imgPath.StartsWith("/"))
                        fullImgPath = imgPath.TrimStart('/');
                    else
                    {
                        var baseUri = new Uri("file:///" + drawingPath);
                        var targetUri = new Uri(baseUri, imgPath);
                        fullImgPath = targetUri.AbsolutePath.TrimStart('/');
                    }
                    
                    var imgEntry = archive.GetEntry(fullImgPath);
                    if (imgEntry != null)
                    {
                        using var imgStream = imgEntry.Open();
                        using var ms = new MemoryStream();
                        imgStream.CopyTo(ms);
                        ws.AddImage(ms.ToArray(), Path.GetExtension(fullImgPath), fromRow, fromCol, toRow - fromRow, toCol - fromCol);
                    }
                }
            }
        }
    }

    private static ExcelFormat DetectFormatFromHeader(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length >= 4 &&
            bytes[0] == 0x50 && bytes[1] == 0x4B &&
            bytes[2] == 0x03 && bytes[3] == 0x04)
        {
            return ExcelFormat.OpenXmlZip;
        }

        if (bytes.Length >= 8 &&
            bytes[0] == 0xD0 && bytes[1] == 0xCF &&
            bytes[2] == 0x11 && bytes[3] == 0xE0 &&
            bytes[4] == 0xA1 && bytes[5] == 0xB1 &&
            bytes[6] == 0x1A && bytes[7] == 0xE1)
        {
            return ExcelFormat.Cfb;
        }

        return ExcelFormat.Unknown;
    }

    public enum ExcelFormat
    {
        Unknown = 0,
        OpenXmlZip = 1,
        Cfb = 2,
    }

    #region XML Generation Helpers (写 Excel 用的辅助方法)

    private readonly record struct DrawingImagePart(string RelationshipId, string MediaFileName, XlWorksheetImage Image, int PictureId);
    private readonly record struct RichValuePart(int RichValueIndex, int ImageIndex);
    private readonly record struct RichValueRelPart(string RelationshipId, string MediaFileName);

    private static string GenerateMetadataXml(int cellImageCount)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xlrd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\">");
        sb.Append("<metadataTypes count=\"1\"><metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\"/></metadataTypes>");
        
        sb.Append("<futureMetadata name=\"XLRICHVALUE\" count=\"" + cellImageCount + "\">");
        for (int i = 0; i < cellImageCount; i++)
        {
            sb.Append("<bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\"><xlrd:rvb i=\"" + i + "\"/></ext></extLst></bk>");
        }
        sb.Append("</futureMetadata>");

        sb.Append("<valueMetadata count=\"" + cellImageCount + "\">");
        for (int i = 0; i < cellImageCount; i++)
        {
            sb.Append("<bk><rc t=\"1\" v=\"" + i + "\"/></bk>");
        }
        sb.Append("</valueMetadata>");
        sb.Append("</metadata>");
        return sb.ToString();
    }

    private static string GenerateRichValueStructureXml()
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               "<rvStructures xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"1\">" +
               "<s t=\"_localImage\"><k n=\"_rvRel:LocalImageIdentifier\" t=\"i\"/><k n=\"CalcOrigin\" t=\"i\"/></s>" +
               "</rvStructures>";
    }

    private static string GenerateRichValueXml(IReadOnlyList<RichValuePart> parts)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<rvData xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"" + parts.Count + "\">");
        foreach (var part in parts)
        {
            // s="0" refers to the first structure (localImage)
            // first <v> is LocalImageIdentifier, second <v> is CalcOrigin (5 = embedded)
            sb.Append("<rv s=\"0\"><v>" + part.ImageIndex + "</v><v>5</v></rv>");
        }
        sb.Append("</rvData>");
        return sb.ToString();
    }

    private static string GenerateRichValueRelXml(IReadOnlyList<RichValueRelPart> parts)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<richValueRels xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        foreach (var part in parts)
        {
            sb.Append("<rel r:id=\"" + part.RelationshipId + "\"/>");
        }
        sb.Append("</richValueRels>");
        return sb.ToString();
    }

    private static string GenerateRichValueRelRels(IReadOnlyList<RichValueRelPart> parts)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"" + NsPkgRel + "\">");
        foreach (var part in parts)
        {
            sb.Append("<Relationship Id=\"" + part.RelationshipId + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/" + part.MediaFileName + "\"/>");
        }
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    private static string GenerateRichValueTypesXml()
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               "<rvTypesInfo xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" xmlns:x=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\">" +
               "<rvTypeInfo name=\"XLRICHVALUE\">" +
               "<keyFlags><key name=\"_rvRel:LocalImageIdentifier\" flags=\"1\"/></keyFlags>" +
               "</rvTypeInfo>" +
               "</rvTypesInfo>";
    }

    private static void WriteEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        writer.Write(content);
    }

    private static void WriteBinaryEntry(ZipArchive archive, string path, byte[] content)
    {
        var entry = archive.CreateEntry(path);
        using var stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }

    private static string GenerateContentTypes(int sheetCount, int drawingCount, bool hasCellImages, IEnumerable<string> imageExtensions)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        foreach (var ext in imageExtensions.Select(x => x.ToLowerInvariant()).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.Ordinal))
        {
            sb.Append("<Default Extension=\"" + ext + "\" ContentType=\"" + GetImageContentType(ext) + "\"/>");
        }
        sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"); // 即使为空也需要
        for (int i = 1; i <= sheetCount; i++)
        {
            sb.Append("<Override PartName=\"/xl/worksheets/sheet" + i + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        }
        for (int i = 1; i <= drawingCount; i++)
        {
            sb.Append("<Override PartName=\"/xl/drawings/drawing" + i + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
        }
        if (hasCellImages)
        {
            sb.Append("<Override PartName=\"/xl/metadata.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml\"/>");
            sb.Append("<Override PartName=\"/xl/richData/rdrichvalue.xml\" ContentType=\"application/vnd.ms-excel.rdrichvalue+xml\"/>");
            sb.Append("<Override PartName=\"/xl/richData/rdrichvaluestructure.xml\" ContentType=\"application/vnd.ms-excel.rdrichvaluestructure+xml\"/>");
            sb.Append("<Override PartName=\"/xl/richData/richValueRel.xml\" ContentType=\"application/vnd.ms-excel.richvaluerel+xml\"/>");
        }
        sb.Append("</Types>");
        return sb.ToString();
    }

    private static string GenerateRootRels()
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
               "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
               "</Relationships>";
    }

    private static string GenerateWorkbookXml(List<XlWorksheet> sheets)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<workbook xmlns=\"{NsMain}\" xmlns:r=\"{NsRel}\">");
        sb.Append("<sheets>");
        for (int i = 0; i < sheets.Count; i++)
        {
            // r:id 必须匹配 workbook.xml.rels 里的 Id
            sb.Append($"<sheet name=\"{sheets[i].Name}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
        }
        sb.Append("</sheets>");
        sb.Append("</workbook>");
        return sb.ToString();
    }

    private static string GenerateWorkbookRels(int sheetCount, bool hasCellImages)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"" + NsPkgRel + "\">");
        for (int i = 1; i <= sheetCount; i++)
        {
            sb.Append("<Relationship Id=\"rId" + i + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet" + i + ".xml\"/>");
        }
        // 添加 Styles 关系
        int nextId = sheetCount + 1;
        sb.Append("<Relationship Id=\"rId" + (nextId++) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        
        if (hasCellImages)
        {
            sb.Append("<Relationship Id=\"rId" + (nextId++) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata\" Target=\"metadata.xml\"/>");
            sb.Append("<Relationship Id=\"rId" + (nextId++) + "\" Type=\"http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue\" Target=\"richData/rdrichvalue.xml\"/>");
            sb.Append("<Relationship Id=\"rId" + (nextId++) + "\" Type=\"http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure\" Target=\"richData/rdrichvaluestructure.xml\"/>");
            sb.Append("<Relationship Id=\"rId" + (nextId++) + "\" Type=\"http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes\" Target=\"richData/rdRichValueTypes.xml\"/>");
            sb.Append("<Relationship Id=\"rId" + (nextId++) + "\" Type=\"http://schemas.microsoft.com/office/2022/10/relationships/richValueRel\" Target=\"richData/richValueRel.xml\"/>");
        }

        sb.Append("</Relationships>");
        return sb.ToString();
    }

    private static string GenerateMinimalStyles()
    {
        // 返回一个最小化的样式 XML，否则 Excel 打开时会报“文件已损坏”
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               $"<styleSheet xmlns=\"{NsMain}\">" +
               "<numFmts count=\"0\"/>" +
               "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>" +
               "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
               "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
               "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
               "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>" +
               "</styleSheet>";
    }

    private static string GenerateWorksheetRels(int drawingIndex)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               $"<Relationships xmlns=\"{NsPkgRel}\">" +
               $"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing{drawingIndex}.xml\"/>" +
               "</Relationships>";
    }

    private static string GenerateDrawingXml(IReadOnlyList<DrawingImagePart> drawingImageParts)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"{NsRel}\">");

        foreach (var part in drawingImageParts)
        {
            int fromColumn = part.Image.ColumnIndex;
            int fromRow = part.Image.RowIndex;
            int toColumn = fromColumn + part.Image.ColumnSpan;
            int toRow = fromRow + part.Image.RowSpan;

            sb.Append("<xdr:twoCellAnchor editAs=\"twoCell\">");
            sb.Append($"<xdr:from><xdr:col>{fromColumn}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>");
            sb.Append($"<xdr:to><xdr:col>{toColumn}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>");
            sb.Append("<xdr:pic>");
            sb.Append($"<xdr:nvPicPr><xdr:cNvPr id=\"{part.PictureId}\" name=\"Picture {part.PictureId}\"/><xdr:cNvPicPr/></xdr:nvPicPr>");
            sb.Append($"<xdr:blipFill><a:blip r:embed=\"{part.RelationshipId}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>");
            sb.Append("<xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>");
            sb.Append("</xdr:pic>");
            sb.Append("<xdr:clientData/>");
            sb.Append("</xdr:twoCellAnchor>");
        }

        sb.Append("</xdr:wsDr>");
        return sb.ToString();
    }

    private static string GenerateDrawingRels(IReadOnlyList<DrawingImagePart> drawingImageParts)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<Relationships xmlns=\"{NsPkgRel}\">");
        foreach (var part in drawingImageParts)
        {
            sb.Append($"<Relationship Id=\"{part.RelationshipId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/{part.MediaFileName}\"/>");
        }
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    private static string GetImageContentType(string extension)
    {
        return extension switch
        {
            "png" => "image/png",
            "jpg" or "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tif" or "tiff" => "image/tiff",
            _ => throw new NotSupportedException($"Unsupported image format: .{extension}")
        };
    }

    private static string GenerateSheetXml(XlWorksheet sheet, XlsxStyleBuilder styleBuilder, Dictionary<(int Row, int Col), int> cellVmMap, string? drawingRelationshipId = null)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        var namespaces = string.IsNullOrEmpty(drawingRelationshipId) ? "xmlns=\"" + NsMain + "\"" : "xmlns=\"" + NsMain + "\" xmlns:r=\"" + NsRel + "\"";
        sb.Append("<worksheet " + namespaces + ">");

        // Sheet Properties (Tab Color)
        if (!string.IsNullOrEmpty(sheet.Options.TabColor))
        {
            sb.Append("<sheetPr><tabColor rgb=\"" + sheet.Options.TabColor + "\"/></sheetPr>");
        }

        // Sheet Views (Gridlines)
        sb.Append("<sheetViews>");
        sb.Append("<sheetView tabSelected=\"1\" workbookViewId=\"0\" showGridLines=\"" + (sheet.Options.ShowGridLines ? "1" : "0") + "\">");
        sb.Append("</sheetView></sheetViews>");

        // Sheet Format (Default Row Height)
        if (sheet.Options.DefaultRowHeight.HasValue)
        {
            sb.Append("<sheetFormatPr defaultRowHeight=\"" + sheet.Options.DefaultRowHeight + "\" customHeight=\"1\"/>");
        }
        else
        {
            sb.Append("<sheetFormatPr defaultRowHeight=\"15\"/>");
        }

        // Columns
        if (sheet.Columns.Count > 0)
        {
            sb.Append("<cols>");
            foreach (var colPair in sheet.Columns.OrderBy(x => x.Key))
            {
                int colIdx = colPair.Key + 1;
                var col = colPair.Value;
                sb.Append("<col min=\"" + colIdx + "\" max=\"" + colIdx + "\"");
                if (col.Width.HasValue) sb.Append(" width=\"" + col.Width + "\" customWidth=\"1\"");
                if (col.Hidden) sb.Append(" hidden=\"1\"");
                if (col.Style != null)
                {
                    int styleIdx = styleBuilder.GetStyleIndex(col.Style);
                    sb.Append(" style=\"" + styleIdx + "\"");
                }
                sb.Append("/>");
            }
            sb.Append("</cols>");
        }

        sb.Append("<sheetData>");

        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            sb.Append("<row r=\"" + (r + 1) + "\"");
            if (row.Height.HasValue) sb.Append(" ht=\"" + row.Height + "\" customHeight=\"1\"");
            if (row.Hidden) sb.Append(" hidden=\"1\"");
            
            if (row.Style != null)
            {
                int rowStyleIdx = styleBuilder.GetStyleIndex(row.Style);
                sb.Append(" s=\"" + rowStyleIdx + "\" customFormat=\"1\"");
            }
            sb.Append(">");

            for (int c = 0; c < row.Cells.Count; c++)
            {
                var cell = row.Cells[c];
                var val = cell.Value;
                
                // Get Style Index
                XlStyle? finalStyle = cell.Style ?? row.Style;
                if (finalStyle == null && sheet.Columns.TryGetValue(c, out var col))
                {
                    finalStyle = col.Style;
                }
                int styleIdx = styleBuilder.GetStyleIndex(finalStyle);

                string colRef = GetColumnName(c) + (r + 1);

                if (cellVmMap.TryGetValue((r, c), out var vmIndex))
                {
                    // Modern Standard: t="e", vm="X", value="#VALUE!"
                    sb.Append("<c r=\"" + colRef + "\" s=\"" + styleIdx + "\" t=\"e\" vm=\"" + vmIndex + "\">");
                    sb.Append("<v>#VALUE!</v>");
                    sb.Append("</c>");
                }
                else
                {
                    if (string.IsNullOrEmpty(val) && styleIdx == 0) continue;

                    sb.Append("<c r=\"" + colRef + "\" s=\"" + styleIdx + "\"");

                    if (!string.IsNullOrEmpty(val))
                    {
                        sb.Append(" t=\"inlineStr\">");
                        sb.Append("<is><t>" + EscapeXml(val) + "</t></is>");
                        sb.Append("</c>");
                    }
                    else
                    {
                        sb.Append("/>");
                    }
                }
            }
            sb.Append("</row>");
        }

        sb.Append("</sheetData>");

        if (sheet.MergedCells.Count > 0)
        {
            sb.Append($"<mergeCells count=\"{sheet.MergedCells.Count}\">");
            foreach (var mergeRef in sheet.MergedCells)
            {
                sb.Append($"<mergeCell ref=\"{mergeRef}\"/>");
            }
            sb.Append("</mergeCells>");
        }

        if (!string.IsNullOrEmpty(drawingRelationshipId))
        {
            sb.Append($"<drawing r:id=\"{drawingRelationshipId}\"/>");
        }
        sb.Append("</worksheet>");
        return sb.ToString();
    }

    private static string GetColumnName(int index)
    {
        string columnName = "";
        while (index >= 0)
        {
            columnName = (char)('A' + (index % 26)) + columnName;
            index = (index / 26) - 1;
        }
        return columnName;
    }

    private static string EscapeXml(string txt)
    {
        if (string.IsNullOrEmpty(txt)) return "";
        return txt.Replace("&", "&amp;")
                  .Replace("<", "&lt;")
                  .Replace(">", "&gt;")
                  .Replace("\"", "&quot;")
                  .Replace("'", "&apos;");
    }

    #endregion
}
