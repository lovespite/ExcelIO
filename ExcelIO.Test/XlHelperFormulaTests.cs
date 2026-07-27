using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ExcelIO.Test;

public class XlHelperFormulaTests
{
    [Fact]
    public void Cell_SetFormula_SetsFormulaAndCachedValue()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var cell = row.Insert(0, "10");

        cell.SetFormula("=A1+A2", "10");

        Assert.Equal("=A1+A2", cell.Formula);
        Assert.Equal("10", cell.Value);
        Assert.True(cell.HasFormula);
    }

    [Fact]
    public void Cell_SetValue_ClearsFormula()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        var row = ws.NewRow();
        var cell = row.Insert(0, "");
        cell.SetFormula("=SUM(A1:A2)", "15");

        cell.Value = "20";

        Assert.Null(cell.Formula);
        Assert.False(cell.HasFormula);
        Assert.Equal("20", cell.Value);
    }

    [Fact]
    public void Save_WithFormula_EmitsFormulaElement()
    {
        var path = Path.Combine(Path.GetTempPath(), $"formula-{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = new XlWorkbook();
            var ws = wb.NewWorksheet("Sheet1");
            var row = ws.NewRow();
            row.Add("5");
            row.Add("3");
            var formulaCell = row.Insert(2, "");
            formulaCell.SetFormula("=A1+B1", "8");

            XlHelper.Save(path, wb);

            using var archive = ZipFile.OpenRead(path);
            var sheetXml = ReadEntryText(archive, "xl/worksheets/sheet1.xml");

            Assert.Contains("<f>=A1+B1</f>", sheetXml);
            Assert.Contains("<v>8</v>", sheetXml);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Load_WithFormula_ParsesFormulaAndValue()
    {
        var xlsxBytes = BuildMinimalXlsxWithFormula("=SUM(A1:A2)", "15");
        using var ms = new MemoryStream(xlsxBytes);
        var wb = XlHelper.Load(ms);

        var ws = wb.Worksheets[0];
        Assert.Single(ws.Rows);
        var cell = ws.Rows[0].Cells[0];

        Assert.Equal("=SUM(A1:A2)", cell.Formula);
        Assert.Equal("15", cell.Value);
        Assert.True(cell.HasFormula);
    }

    [Fact]
    public void RoundTrip_FormulasPreserved()
    {
        var path1 = Path.Combine(Path.GetTempPath(), $"roundtrip1-{Guid.NewGuid():N}.xlsx");
        var path2 = Path.Combine(Path.GetTempPath(), $"roundtrip2-{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = new XlWorkbook();
            var ws = wb.NewWorksheet("Sheet1");
            var row = ws.NewRow();
            row.Add("10");
            row.Add("20");
            var sumCell = row.Insert(2, "");
            sumCell.SetFormula("=A1+B1", "30");

            XlHelper.Save(path1, wb);

            var wb2 = XlHelper.Load(path1);
            XlHelper.Save(path2, wb2);

            var wb3 = XlHelper.Load(path2);
            var cell = wb3.Worksheets[0].Rows[0].Cells[2];

            Assert.Equal("=A1+B1", cell.Formula);
            Assert.Equal("30", cell.Value);
        }
        finally
        {
            if (File.Exists(path1)) File.Delete(path1);
            if (File.Exists(path2)) File.Delete(path2);
        }
    }

    [Fact]
    public void Load_SharedFormula_Master_ParsesFormula()
    {
        var xlsxBytes = BuildXlsxWithSharedFormula();
        using var ms = new MemoryStream(xlsxBytes);
        var wb = XlHelper.Load(ms);

        var ws = wb.Worksheets[0];
        Assert.Equal(3, ws.Rows.Count);

        // Master formula cell (A1)
        var masterCell = ws.Rows[0].Cells[0];
        Assert.Equal("=B1*2", masterCell.Formula);
        Assert.Equal("20", masterCell.Value);
    }

    [Fact]
    public void Load_SharedFormula_ChildCells_TranslateReferences()
    {
        var xlsxBytes = BuildXlsxWithSharedFormula();
        using var ms = new MemoryStream(xlsxBytes);
        var wb = XlHelper.Load(ms);

        var ws = wb.Worksheets[0];

        // Child cells should have translated formulas
        var cell2 = ws.Rows[1].Cells[0]; // A2
        Assert.Equal("=B2*2", cell2.Formula);
        Assert.Equal("40", cell2.Value);

        var cell3 = ws.Rows[2].Cells[0]; // A3
        Assert.Equal("=B3*2", cell3.Formula);
        Assert.Equal("60", cell3.Value);
    }

    [Fact]
    public void Load_SharedFormula_WithAbsoluteReferences_PreservesAbsolute()
    {
        var xlsxBytes = BuildXlsxWithAbsoluteSharedFormula();
        using var ms = new MemoryStream(xlsxBytes);
        var wb = XlHelper.Load(ms);

        var ws = wb.Worksheets[0];

        // Master: =$B$1*2
        var master = ws.Rows[0].Cells[0];
        Assert.Equal("=$B$1*2", master.Formula);

        // Child cells: $B$1 should not change
        var child = ws.Rows[1].Cells[0];
        Assert.Equal("=$B$1*2", child.Formula);
    }

    [Fact]
    public void Load_SharedFormula_MixedReferences_TranslatesCorrectly()
    {
        var xlsxBytes = BuildXlsxWithMixedRefSharedFormula();
        using var ms = new MemoryStream(xlsxBytes);
        var wb = XlHelper.Load(ms);

        var ws = wb.Worksheets[0];

        // Master: =B1+$C$1
        var master = ws.Rows[0].Cells[0];
        Assert.Equal("=B1+$C$1", master.Formula);

        // Child: B should shift, C should not
        var child = ws.Rows[1].Cells[0];
        Assert.Equal("=B2+$C$1", child.Formula);
    }

    // ── Helper methods ──

    private static string ReadEntryText(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName);
        if (entry == null) return "";
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static byte[] BuildMinimalXlsxWithFormula(string formula, string cachedValue)
    {
        var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "[Content_Types].xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Default Extension=""xml"" ContentType=""application/xml""/>
<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
</Types>");

            WriteEntry(archive, "_rels/.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>");

            WriteEntry(archive, "xl/_rels/workbook.xml.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
</Relationships>");

            WriteEntry(archive, "xl/workbook.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets><sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1""/></sheets>
</workbook>");

            var sheetXml = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<sheetData>
<row r=""1"">
<c r=""A1""><f>{System.Security.SecurityElement.Escape(formula)}</f><v>{System.Security.SecurityElement.Escape(cachedValue)}</v></c>
</row>
</sheetData>
</worksheet>";
            WriteEntry(archive, "xl/worksheets/sheet1.xml", sheetXml);
        }
        return ms.ToArray();
    }

    private static byte[] BuildXlsxWithSharedFormula()
    {
        var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "[Content_Types].xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Default Extension=""xml"" ContentType=""application/xml""/>
<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
</Types>");

            WriteEntry(archive, "_rels/.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>");

            WriteEntry(archive, "xl/_rels/workbook.xml.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
</Relationships>");

            WriteEntry(archive, "xl/workbook.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets><sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1""/></sheets>
</workbook>");

            // Shared formula: master in A1:A3, formula =B1*2
            var sheetXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<sheetData>
<row r=""1"">
<c r=""A1""><f t=""shared"" ref=""A1:A3"" si=""0"">=B1*2</f><v>20</v></c>
<c r=""B1""><v>10</v></c>
</row>
<row r=""2"">
<c r=""A2""><f t=""shared"" si=""0""/><v>40</v></c>
<c r=""B2""><v>20</v></c>
</row>
<row r=""3"">
<c r=""A3""><f t=""shared"" si=""0""/><v>60</v></c>
<c r=""B3""><v>30</v></c>
</row>
</sheetData>
</worksheet>";
            WriteEntry(archive, "xl/worksheets/sheet1.xml", sheetXml);
        }
        return ms.ToArray();
    }

    private static byte[] BuildXlsxWithAbsoluteSharedFormula()
    {
        var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "[Content_Types].xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Default Extension=""xml"" ContentType=""application/xml""/>
<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
</Types>");

            WriteEntry(archive, "_rels/.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>");

            WriteEntry(archive, "xl/_rels/workbook.xml.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
</Relationships>");

            WriteEntry(archive, "xl/workbook.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets><sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1""/></sheets>
</workbook>");

            var sheetXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<sheetData>
<row r=""1"">
<c r=""A1""><f t=""shared"" ref=""A1:A2"" si=""0"">=$B$1*2</f><v>20</v></c>
<c r=""B1""><v>10</v></c>
</row>
<row r=""2"">
<c r=""A2""><f t=""shared"" si=""0""/><v>20</v></c>
</row>
</sheetData>
</worksheet>";
            WriteEntry(archive, "xl/worksheets/sheet1.xml", sheetXml);
        }
        return ms.ToArray();
    }

    private static byte[] BuildXlsxWithMixedRefSharedFormula()
    {
        var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "[Content_Types].xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Default Extension=""xml"" ContentType=""application/xml""/>
<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
</Types>");

            WriteEntry(archive, "_rels/.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>");

            WriteEntry(archive, "xl/_rels/workbook.xml.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
</Relationships>");

            WriteEntry(archive, "xl/workbook.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets><sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1""/></sheets>
</workbook>");

            var sheetXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<sheetData>
<row r=""1"">
<c r=""A1""><f t=""shared"" ref=""A1:A2"" si=""0"">=B1+$C$1</f><v>15</v></c>
<c r=""B1""><v>10</v></c>
<c r=""C1""><v>5</v></c>
</row>
<row r=""2"">
<c r=""A2""><f t=""shared"" si=""0""/><v>25</v></c>
<c r=""B2""><v>20</v></c>
</row>
</sheetData>
</worksheet>";
            WriteEntry(archive, "xl/worksheets/sheet1.xml", sheetXml);
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string entryName, string content)
    {
        var entry = archive.CreateEntry(entryName);
        using var stream = entry.Open();
        var bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }
}
