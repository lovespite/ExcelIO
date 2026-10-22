using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using ExcelIO;

namespace QuickExcel.Core;

internal static class DocumentPackage
{
    private const long MaxFileBytes = 64L * 1024 * 1024;
    private const long MaxExpandedBytes = 256L * 1024 * 1024;
    private const int MaxScriptBytes = 1024 * 1024;
    private static readonly UTF8Encoding Utf8 = new(false, true);

    public static (XlWorkbook Workbook, string Script) Open(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var extension = Path.GetExtension(path);
        if (!extension.Equals(".xlsq", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("原型支持打开 .xlsx 和 .xlsq 文件。");
        using var input = File.OpenRead(path);
        if (input.Length > MaxFileBytes) throw new InvalidDataException("原型支持的文件大小上限为 64 MiB。");
        if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            return (LoadWorkbook(input), WorkbookDocument.DefaultScript);

        using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);
        ValidateArchive(archive, maxEntries: 16);
        using var manifest = JsonDocument.Parse(ReadEntry(archive, "manifest.json", 64 * 1024));
        if (!manifest.RootElement.TryGetProperty("formatVersion", out var version) || !version.TryGetInt32(out var versionNumber) || versionNumber != 1)
            throw new InvalidDataException("不支持此 XLSQ 文件版本，当前原型仅支持 formatVersion 1。");
        if (manifest.RootElement.TryGetProperty("format", out var format) && format.GetString() != "QuickExcel")
            throw new InvalidDataException("文件不是 QuickExcel XLSQ 工作簿。");
        var script = Utf8.GetString(ReadEntry(archive, "scripts/functions.js", MaxScriptBytes));
        using var workbookStream = new MemoryStream(ReadEntry(archive, "workbook.xlsx", MaxFileBytes), writable: false);
        return (LoadWorkbook(workbookStream), script);
    }

    public static void Save(string path, XlWorkbook workbook, string script)
    {
        RequireExtension(path, ".xlsq");
        var scriptBytes = Utf8.GetBytes(script);
        if (scriptBytes.Length > MaxScriptBytes) throw new InvalidOperationException("脚本大小不能超过 1 MiB。");
        // A copy also prevents the existing XLSX writer's image-anchor expansion from changing the editor model.
        var snapshot = Clone(workbook, valuesOnly: false);
        using var workbookStream = new MemoryStream();
        Invariant(() => { XlHelper.Save(workbookStream, snapshot); return true; });
        if (workbookStream.Length > MaxFileBytes) throw new InvalidOperationException("工作簿超过原型支持的 64 MiB 上限。");
        AtomicWrite(path, output =>
        {
            using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);
            WriteEntry(archive, "manifest.json", JsonSerializer.SerializeToUtf8Bytes(new
            {
                format = "QuickExcel",
                formatVersion = 1,
                scriptApiVersion = 1,
                workbook = "workbook.xlsx",
                script = "scripts/functions.js"
            }, new JsonSerializerOptions { WriteIndented = true }));
            WriteEntry(archive, "workbook.xlsx", workbookStream.ToArray());
            WriteEntry(archive, "scripts/functions.js", scriptBytes);
        });
    }

    public static void Export(string path, XlWorkbook workbook)
    {
        RequireExtension(path, ".xlsx");
        var snapshot = Clone(workbook, valuesOnly: true);
        AtomicWrite(path, output => Invariant(() => { XlHelper.Save(output, snapshot); return true; }));
    }

    private static void RequireExtension(string path, string extension)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!Path.GetExtension(path).Equals(extension, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException($"文件名必须以 {extension} 结尾。", nameof(path));
    }

    private static void AtomicWrite(string path, Action<Stream> write)
    {
        var fullPath = Path.GetFullPath(path);
        var temporaryPath = Path.Combine(Path.GetDirectoryName(fullPath)!, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (var output = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                write(output);
                if (output.Length > MaxFileBytes) throw new InvalidOperationException("文件超过原型支持的 64 MiB 上限。");
                output.Flush(flushToDisk: true);
            }
            if (File.Exists(fullPath)) File.Replace(temporaryPath, fullPath, destinationBackupFileName: null);
            else File.Move(temporaryPath, fullPath);
        }
        finally
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    private static XlWorkbook LoadWorkbook(Stream stream)
    {
        List<Dictionary<int, List<int>>> layouts;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true))
        {
            ValidateArchive(archive, maxEntries: 2048);
            layouts = ValidateWorksheetDimensions(archive);
        }
        stream.Position = 0;
        var workbook = Invariant(() => XlHelper.Load(stream));
        // ExcelIO currently appends cells instead of positioning them by the XLSX r attribute.
        // Restore sparse coordinates here so an empty A1 cannot shift B1 into the wrong column.
        for (var sheetIndex = 0; sheetIndex < Math.Min(workbook.Count, layouts.Count); sheetIndex++)
        foreach (var (rowIndex, columns) in layouts[sheetIndex])
        {
            var row = workbook[sheetIndex].Rows[rowIndex];
            if (columns.Count != row.Cells.Count) throw new InvalidDataException("工作表单元格布局不一致。");
            var cells = new List<XlCell>();
            for (var index = 0; index < columns.Count; index++)
            {
                while (cells.Count < columns[index]) cells.Add(new XlCell(row));
                cells.Add(row.Cells[index]);
            }
            row.Cells = cells;
        }
        return workbook;
    }

    private static void ValidateArchive(ZipArchive archive, int maxEntries)
    {
        if (archive.Entries.Count > maxEntries) throw new InvalidDataException("压缩包包含过多条目。");
        var names = new HashSet<string>(StringComparer.Ordinal);
        long total = 0;
        foreach (var entry in archive.Entries)
        {
            if (!names.Add(entry.FullName)) throw new InvalidDataException($"压缩包包含重复条目：{entry.FullName}");
            if (entry.Length > MaxFileBytes || entry.Length > MaxExpandedBytes - total)
                throw new InvalidDataException("压缩包展开后超过原型容量限制。");
            total += entry.Length;
        }
    }

    private static T Invariant<T>(Func<T> action)
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            return action();
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    private static byte[] ReadEntry(ZipArchive archive, string name, long limit)
    {
        var entry = archive.GetEntry(name) ?? throw new InvalidDataException($"XLSQ 文件缺少 {name}。");
        if (entry.Length > limit) throw new InvalidDataException($"{name} 超过容量限制。");
        using var input = entry.Open();
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        int count;
        while ((count = input.Read(buffer)) > 0)
        {
            if (output.Length + count > limit) throw new InvalidDataException($"{name} 超过容量限制。");
            output.Write(buffer, 0, count);
        }
        return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, byte[] bytes)
    {
        using var output = archive.CreateEntry(name, CompressionLevel.Optimal).Open();
        output.Write(bytes);
    }

    private static List<Dictionary<int, List<int>>> ValidateWorksheetDimensions(ZipArchive archive)
    {
        long totalCells = 0;
        var layouts = new List<Dictionary<int, List<int>>>();
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaxFileBytes };
        foreach (var entry in ResolveWorksheets(archive, settings))
        {
            using var input = entry.Open();
            using var reader = XmlReader.Create(input, settings);
            var rowWidths = new Dictionary<int, int>();
            var layout = new Dictionary<int, List<int>>();
            layouts.Add(layout);
            var currentRow = 0;
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "row")
                {
                    var index = reader.GetAttribute("r");
                    var nextRow = index is null ? 0 : ParsePositive(index, WorkbookDocument.MaxRows);
                    if (nextRow <= currentRow) throw new InvalidDataException("工作表的行地址缺失、重复或未按顺序排列。");
                    currentRow = nextRow;
                    if (currentRow > WorkbookDocument.MaxRows) throw new InvalidDataException("工作表行数超过原型限制。");
                    layout[currentRow - 1] = [];
                }
                else if (reader.LocalName == "c")
                {
                    var reference = reader.GetAttribute("r");
                    var column = 0;
                    if (reference is not null)
                    {
                        var position = 0;
                        while (position < reference.Length && reference[position] is >= 'A' and <= 'Z')
                        {
                            column = checked(column * 26 + reference[position++] - 'A' + 1);
                            if (column > WorkbookDocument.MaxColumns) throw new InvalidDataException("工作表列数超过原型限制。");
                        }
                        if (position == 0 || position == reference.Length) throw new InvalidDataException("无效的单元格地址。");
                        if (ParsePositive(reference[position..], WorkbookDocument.MaxRows) != currentRow)
                            throw new InvalidDataException("单元格地址与所在行不一致。");
                    }
                    else column = rowWidths.GetValueOrDefault(currentRow) + 1;
                    var previousWidth = rowWidths.GetValueOrDefault(currentRow);
                    if (currentRow == 0 || column <= previousWidth) throw new InvalidDataException("工作表的单元格地址重复或未按顺序排列。");
                    layout[currentRow - 1].Add(column - 1);
                    if (column > previousWidth)
                    {
                        totalCells += column - previousWidth;
                        rowWidths[currentRow] = column;
                    }
                    if (column > WorkbookDocument.MaxColumns || totalCells > WorkbookDocument.MaxCells)
                        throw new InvalidDataException("工作簿超过原型支持的单元格数量限制。");
                }
                else if (reader.LocalName == "col")
                {
                    if (reader.GetAttribute("max") is { } maximum) _ = ParsePositive(maximum, WorkbookDocument.MaxColumns);
                }
            }
        }
        return layouts;
    }

    private static IEnumerable<ZipArchiveEntry> ResolveWorksheets(ZipArchive archive, XmlReaderSettings settings)
    {
        var workbook = archive.GetEntry("xl/workbook.xml");
        var relationships = archive.GetEntry("xl/_rels/workbook.xml.rels");
        // Standalone worksheet fixtures are still checked before the normal loader rejects them.
        if (workbook is null || relationships is null)
            return archive.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal)).ToArray();
        XDocument ReadXml(ZipArchiveEntry entry)
        {
            using var input = entry.Open();
            using var reader = XmlReader.Create(input, settings);
            return XDocument.Load(reader);
        }
        XNamespace packageNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace relationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var mapping = ReadXml(relationships).Descendants(packageNs + "Relationship")
            .Where(element => element.Attribute("Id") is not null && element.Attribute("Target") is not null)
            .ToDictionary(element => (string)element.Attribute("Id")!, element => (string)element.Attribute("Target")!);
        var result = new List<ZipArchiveEntry>();
        foreach (var sheet in ReadXml(workbook).Descendants(mainNs + "sheet"))
        {
            if ((string?)sheet.Attribute(relationshipNs + "id") is not { } id || !mapping.TryGetValue(id, out var target)) continue;
            var path = target.StartsWith('/') ? target.TrimStart('/') : $"xl/{target}";
            if (!path.StartsWith("xl/", StringComparison.Ordinal)) path = "xl/" + path;
            var entry = archive.GetEntry(path) ?? archive.GetEntry(target);
            if (entry is not null) result.Add(entry);
            if (result.Count > 64) throw new InvalidDataException("原型最多支持 64 张工作表。");
        }
        return result;
    }

    private static int ParsePositive(string value, int maximum)
    {
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var result) || result < 1 || result > maximum)
            throw new InvalidDataException("工作表尺寸无效或超过原型限制。");
        return result;
    }

    private static XlWorkbook Clone(XlWorkbook workbook, bool valuesOnly)
    {
        var clone = new XlWorkbook();
        foreach (var sheet in workbook)
        {
            var target = clone.NewWorksheet(sheet.Name);
            target.Options.ShowGridLines = sheet.Options.ShowGridLines;
            target.Options.DefaultRowHeight = sheet.Options.DefaultRowHeight;
            target.Options.TabColor = sheet.Options.TabColor;
            foreach (var (index, column) in sheet.Columns)
                target.Columns[index] = new XlColumn { Width = column.Width, Hidden = column.Hidden, Style = column.Style };
            target.MergedCells.AddRange(sheet.MergedCells);
            foreach (var row in sheet.Rows)
            {
                var targetRow = target.NewRow();
                targetRow.Height = row.Height;
                targetRow.Hidden = row.Hidden;
                targetRow.Style = row.Style;
                foreach (var cell in row.Cells)
                {
                    var targetCell = new XlCell(targetRow) { Style = cell.Style };
                    // SetFormula(null) is not used: assigning cached text must not accidentally retain formulas.
                    if (!valuesOnly && cell.HasFormula) targetCell.SetFormula(cell.Formula!, cell.Value);
                    else targetCell.SetCalculatedValue(cell.Value);
                    targetRow.Cells.Add(targetCell);
                }
            }
            foreach (var image in sheet.Images)
                target.AddImage(image.Bytes, image.Extension, image.RowIndex, image.ColumnIndex, image.RowSpan, image.ColumnSpan, image.PlaceInCell);
        }
        if (clone.Count == 0) clone.NewWorksheet("Sheet1");
        return clone;
    }
}
