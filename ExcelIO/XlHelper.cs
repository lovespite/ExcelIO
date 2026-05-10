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

    public static Task<XlWorkbook> LoadAsync(string filepath)
        => Task.Run(() => Load(filepath));

    public static async Task<XlWorkbook> LoadAsync(Stream stream, string extension)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return LoadByExtension(stream, extension);
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
        var imageExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int drawingIndex = 1;
        for (int i = 0; i < workbookData.Worksheets.Count; i++)
        {
            var sheet = workbookData.Worksheets[i];
            if (sheet.Images.Count == 0) continue;

            drawingSheets.Add((i + 1, drawingIndex, sheet));
            drawingIndex++;
            foreach (var image in sheet.Images)
            {
                imageExtensions.Add(image.Extension);
            }
        }
        var drawingSheetMap = drawingSheets.ToDictionary(x => x.SheetIndex, x => x.DrawingIndex);

        using var fileStream = new FileStream(filepath, FileMode.Create);
        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Create);

        // 1. 写入 [Content_Types].xml (定义文件类型)
        WriteEntry(archive, "[Content_Types].xml", GenerateContentTypes(workbookData.Worksheets.Count, drawingSheets.Count, imageExtensions));

        // 2. 写入 _rels/.rels (定义根关系)
        WriteEntry(archive, "_rels/.rels", GenerateRootRels());

        // 3. 写入 xl/workbook.xml (定义工作簿结构)
        WriteEntry(archive, "xl/workbook.xml", GenerateWorkbookXml(workbookData.Worksheets));

        // 4. 写入 xl/_rels/workbook.xml.rels (定义工作簿与Sheet的关系)
        WriteEntry(archive, "xl/_rels/workbook.xml.rels", GenerateWorkbookRels(workbookData.Worksheets.Count));

        var styleBuilder = new XlsxStyleBuilder();

        // 6. 写入具体的 Sheet 数据
        for (int i = 0; i < workbookData.Worksheets.Count; i++)
        {
            var sheet = workbookData.Worksheets[i];
            drawingSheetMap.TryGetValue(i + 1, out int sheetDrawingIndex);
            var path = $"xl/worksheets/sheet{i + 1}.xml";
            WriteEntry(archive, path, GenerateSheetXml(sheet, styleBuilder, sheetDrawingIndex > 0 ? "rId1" : null));
            if (sheetDrawingIndex > 0)
            {
                WriteEntry(archive, $"xl/worksheets/_rels/sheet{i + 1}.xml.rels", GenerateWorksheetRels(sheetDrawingIndex));
            }
        }

        // 5. 写入 xl/styles.xml
        WriteEntry(archive, "xl/styles.xml", styleBuilder.GenerateStylesXml());

        int globalImageIndex = 1;
        foreach (var (_, currentDrawingIndex, sheet) in drawingSheets)
        {
            var drawingImageParts = new List<DrawingImagePart>(sheet.Images.Count);
            for (int imageIndex = 0; imageIndex < sheet.Images.Count; imageIndex++)
            {
                var image = sheet.Images[imageIndex];
                var mediaFileName = $"image{globalImageIndex}.{image.Extension}";
                WriteBinaryEntry(archive, $"xl/media/{mediaFileName}", image.Bytes);
                drawingImageParts.Add(new DrawingImagePart($"rId{imageIndex + 1}", mediaFileName, image, imageIndex + 1));
                globalImageIndex++;
            }

            WriteEntry(archive, $"xl/drawings/drawing{currentDrawingIndex}.xml", GenerateDrawingXml(drawingImageParts));
            WriteEntry(archive, $"xl/drawings/_rels/drawing{currentDrawingIndex}.xml.rels", GenerateDrawingRels(drawingImageParts));
        }
    }

    /// <summary>
    /// Loads an Excel workbook from the specified file path, automatically detecting the file format based on the file
    /// extension.
    /// </summary>
    /// <remarks>The method supports both legacy .xls and modern .xlsx/.xlsm formats. The file format is
    /// determined by the file extension. The file is opened with shared read access, allowing other processes to read
    /// or write to the file concurrently.</remarks>
    /// <param name="filepath">The path to the Excel file to load. The file must exist and be accessible for reading. Cannot be null, empty, or
    /// consist only of white-space characters.</param>
    /// <returns>An <see cref="XlWorkbook"/> instance representing the contents of the loaded Excel file.</returns>
    public static XlWorkbook Load(string filepath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filepath);
        using var fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return LoadByExtension(fs, filepath);
    }

    /// <summary>
    /// Loads an Excel workbook from the specified stream, using the provided file extension to determine the file
    /// format.
    /// </summary>
    /// <remarks>The method determines the Excel file format based on the provided extension and parses the
    /// stream accordingly. The caller is responsible for managing the lifetime of the input stream.</remarks>
    /// <param name="stream">The stream containing the Excel file data to load. Cannot be null and must be readable and seekable.</param>
    /// <param name="extension">The file extension that indicates the format of the Excel file (for example, ".xls" or ".xlsx"). Used to select
    /// the appropriate parser.</param>
    /// <returns>An instance of <see cref="XlWorkbook"/> representing the loaded workbook.</returns>
    public static XlWorkbook Load(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return LoadBySignature(stream);
    }

    /// <summary>
    /// Loads an Excel workbook from the specified stream using the given file format.
    /// </summary>
    /// <remarks>The caller is responsible for managing the lifetime of the provided stream. The method
    /// supports both OpenXML (.xlsx, .xlsm) and legacy binary (.xls) Excel formats, as determined by the format
    /// parameter.</remarks>
    /// <param name="stream">The stream containing the Excel file data to load. The stream must be readable and positioned at the start of
    /// the file.</param>
    /// <param name="format">The format of the Excel file to load. Specifies how the stream should be interpreted.</param>
    /// <returns>An XlWorkbook representing the contents of the loaded Excel file.</returns>
    public static XlWorkbook Load(Stream stream, ExcelFormat format)
    {
        return format switch
        {
            ExcelFormat.OpenXmlZip => LoadOpenXml(stream),
            ExcelFormat.Cfb => LoadByExtension(stream, ".xls"),
            _ => LoadBySignature(stream),
        };
    }


    private static XlWorkbook LoadByExtension(Stream stream, string extension)
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

        return LoadOpenXml(stream);
    }

    private static XlWorkbook LoadBySignature(Stream stream)
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
            return Load(stream, format);
        }
        catch (NotSupportedException)
        {
            using var ms = new MemoryStream();
            ms.Write(header, 0, bytesRead);
            stream.CopyTo(ms);
            ms.Seek(0, SeekOrigin.Begin);
            return Load(ms, format);
        }
    }

    private static XlWorkbook LoadOpenXml(Stream stream)
    {
        var result = new XlWorkbook();

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        // 1. 读取 SharedStrings (共享字符串表)
        // Excel 为了压缩体积，把重复的字符串放在一个表里，单元格里存索引
        var sharedStrings = new List<string>();
        var sharedStringEntry = archive.GetEntry("xl/sharedStrings.xml");
        if (sharedStringEntry != null)
        {
            using var entryStream = sharedStringEntry.Open();
            var doc = XDocument.Load(entryStream);
            // <si><t>Value</t></si>
            foreach (var si in doc.Descendants(XName.Get("si", NsMain)))
            {
                // 有时候文本在 <t> 中，有时候在 <r><t> 中 (Rich Text)
                var val = si.DescendantNodes().OfType<XText>().Select(x => x.Value).Aggregate(new StringBuilder(), (sb, v) => sb.Append(v)).ToString();
                sharedStrings.Add(val);
            }
        }

        // 2. 读取 Workbook (获取 Sheet 名称和 ID 的对应关系)
        // 结构: <sheet name="Sheet1" sheetId="1" r:id="rId1" />
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
                var rid = sheet.Attribute(XName.Get("id", NsRel))?.Value; // r:id
                if (rid != null)
                {
                    sheetMapping.Add((name, rid));
                }
            }
        }

        // 3. 解析 Workbook 关系文件，找到 rId 对应的文件名
        // 结构: <Relationship Id="rId1" Target="worksheets/sheet1.xml" />
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

        // 4. 遍历并读取每个 Sheet 的数据
        foreach (var (sheetName, rid) in sheetMapping)
        {
            if (!relMapping.TryGetValue(rid, out var targetPath)) continue;

            // 处理路径差异 (有时候是绝对路径，有时候是相对路径)
            // 简单处理：如果不是以 xl/ 开头，就拼上去
            string entryPath = targetPath.StartsWith("/") ? targetPath.TrimStart('/') : $"xl/{targetPath}";
            // 如果 Target 是 "worksheets/sheet1.xml"，在 zip 里通常是 "xl/worksheets/sheet1.xml"
            if (!entryPath.StartsWith("xl/")) entryPath = "xl/" + entryPath;

            var sheetEntry = archive.GetEntry(entryPath);
            // 容错：有些工具生成的路径可能不同，尝试直接用 Target
            if (sheetEntry == null) sheetEntry = archive.GetEntry(targetPath);

            if (sheetEntry == null) continue;

            var ws = new XlWorksheet(result) { Name = sheetName };
            result.Worksheets.Add(ws);

            using var entryStream = sheetEntry.Open();
            using var reader = XmlReader.Create(entryStream);

            XlRow? currentRow = null;
            int currentRowIndex = -1;

            while (reader.Read())
            {
                // <row r="1">
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
                {
                    var rIndexAttr = reader.GetAttribute("r");
                    if (int.TryParse(rIndexAttr, out int rIndex))
                    {
                        // Excel 索引从 1 开始
                        rIndex -= 1;
                        // 填充空行
                        while (ws.Rows.Count < rIndex)
                        {
                            ws.Rows.Add(new XlRow(ws));
                        }
                        currentRow = new XlRow(ws);
                        ws.Rows.Add(currentRow);
                        currentRowIndex = rIndex;
                    }
                }
                // <c r="A1" t="s">
                else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
                {
                    if (currentRow == null) continue;

                    var type = reader.GetAttribute("t"); // 类型: s=sharedString, str=string, inlineStr=inline

                    // 读取值 <v> 或 <t>
                    // 简单的读取逻辑：读取子树文本
                    // 注意：这里需要根据 type 来判断如何解析
                    string cellValue = string.Empty;

                    // 为了性能，我们手动 Read 到下一个元素
                    // 这是一个简化的 XML 解析，只针对简单的 Excel 结构
                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            if (reader.LocalName == "v") // Value
                            {
                                var raw = reader.ReadElementContentAsString();
                                if (type == "s" && int.TryParse(raw, out int idx) && idx < sharedStrings.Count)
                                {
                                    cellValue = sharedStrings[idx];
                                }
                                else if (type == "b") // boolean
                                {
                                    cellValue = raw == "1" ? "TRUE" : "FALSE";
                                }
                                else
                                {
                                    cellValue = raw;
                                }
                                break;
                            }
                            else if (reader.LocalName == "t") // Inline Text
                            {
                                cellValue = reader.ReadElementContentAsString();
                                break;
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "c")
                        {
                            break;
                        }
                    }

                    currentRow.Cells.Add(new XlCell(currentRow) { Value = cellValue });
                }
            }
        }

        return result;
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

    private static string GenerateContentTypes(int sheetCount, int drawingCount, IEnumerable<string> imageExtensions)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        foreach (var ext in imageExtensions.Select(x => x.ToLowerInvariant()).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.Ordinal))
        {
            sb.Append($"<Default Extension=\"{ext}\" ContentType=\"{GetImageContentType(ext)}\"/>");
        }
        sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"); // 即使为空也需要
        for (int i = 1; i <= sheetCount; i++)
        {
            sb.Append($"<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        }
        for (int i = 1; i <= drawingCount; i++)
        {
            sb.Append($"<Override PartName=\"/xl/drawings/drawing{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
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

    private static string GenerateWorkbookRels(int sheetCount)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<Relationships xmlns=\"{NsPkgRel}\">");
        for (int i = 1; i <= sheetCount; i++)
        {
            sb.Append($"<Relationship Id=\"rId{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>");
        }
        // 添加 Styles 关系
        sb.Append($"<Relationship Id=\"rId{sheetCount + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
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

    private static string GenerateSheetXml(XlWorksheet sheet, XlsxStyleBuilder styleBuilder, string? drawingRelationshipId = null)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        var namespaces = string.IsNullOrEmpty(drawingRelationshipId) ? $"xmlns=\"{NsMain}\"" : $"xmlns=\"{NsMain}\" xmlns:r=\"{NsRel}\"";
        sb.Append($"<worksheet {namespaces}>");

        // Sheet Properties (Tab Color)
        if (!string.IsNullOrEmpty(sheet.Options.TabColor))
        {
            sb.Append($"<sheetPr><tabColor rgb=\"{sheet.Options.TabColor}\"/></sheetPr>");
        }

        // Sheet Views (Gridlines)
        sb.Append("<sheetViews>");
        sb.Append($"<sheetView tabSelected=\"1\" workbookViewId=\"0\" showGridLines=\"{(sheet.Options.ShowGridLines ? "1" : "0")}\">");
        sb.Append("</sheetView></sheetViews>");

        // Sheet Format (Default Row Height)
        if (sheet.Options.DefaultRowHeight.HasValue)
        {
            sb.Append($"<sheetFormatPr defaultRowHeight=\"{sheet.Options.DefaultRowHeight}\" customHeight=\"1\"/>");
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
                sb.Append($"<col min=\"{colIdx}\" max=\"{colIdx}\"");
                if (col.Width.HasValue) sb.Append($" width=\"{col.Width}\" customWidth=\"1\"");
                if (col.Hidden) sb.Append(" hidden=\"1\"");
                if (col.Style != null)
                {
                    int styleIdx = styleBuilder.GetStyleIndex(col.Style);
                    sb.Append($" style=\"{styleIdx}\"");
                }
                sb.Append("/>");
            }
            sb.Append("</cols>");
        }

        sb.Append("<sheetData>");

        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            sb.Append($"<row r=\"{r + 1}\"");
            if (row.Height.HasValue) sb.Append($" ht=\"{row.Height}\" customHeight=\"1\"");
            if (row.Hidden) sb.Append(" hidden=\"1\"");
            
            if (row.Style != null)
            {
                int rowStyleIdx = styleBuilder.GetStyleIndex(row.Style);
                sb.Append($" s=\"{rowStyleIdx}\" customFormat=\"1\"");
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

                if (string.IsNullOrEmpty(val) && styleIdx == 0) continue;

                string colRef = GetColumnName(c) + (r + 1);
                sb.Append($"<c r=\"{colRef}\" s=\"{styleIdx}\"");

                if (!string.IsNullOrEmpty(val))
                {
                    sb.Append(" t=\"inlineStr\">");
                    sb.Append($"<is><t>{EscapeXml(val)}</t></is>");
                    sb.Append("</c>");
                }
                else
                {
                    sb.Append("/>");
                }
            }
            sb.Append("</row>");
        }

        sb.Append("</sheetData>");
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
