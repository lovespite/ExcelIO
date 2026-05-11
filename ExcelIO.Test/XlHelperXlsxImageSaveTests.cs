using System.Diagnostics;
using System.IO.Compression;

namespace ExcelIO.Test;

public class XlHelperXlsxImageSaveTests
{
    private static readonly byte[] Png1x1Bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ioAAAAASUVORK5CYII=");

    [Fact]
    public void Save_Xlsx_WithImageBytes_WritesDrawingAndMediaParts()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-img-{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = new XlWorkbook();
            var ws = wb.NewWorksheet("Sheet1");
            ws.AddRow("header");
            ws.AddImage(Png1x1Bytes, "png", rowIndex: 0, columnIndex: 0);

            XlHelper.Save(path, wb);

            using var archive = ZipFile.OpenRead(path);
            Assert.NotNull(archive.GetEntry("xl/media/image1.png"));
            Assert.NotNull(archive.GetEntry("xl/drawings/drawing1.xml"));
            Assert.NotNull(archive.GetEntry("xl/drawings/_rels/drawing1.xml.rels"));
            Assert.NotNull(archive.GetEntry("xl/worksheets/_rels/sheet1.xml.rels"));

            var sheetXml = ReadEntryText(archive, "xl/worksheets/sheet1.xml");
            var drawingXml = ReadEntryText(archive, "xl/drawings/drawing1.xml");
            var drawingRelsXml = ReadEntryText(archive, "xl/drawings/_rels/drawing1.xml.rels");
            var contentTypesXml = ReadEntryText(archive, "[Content_Types].xml");

            Assert.Contains("<drawing r:id=\"rId1\"/>", sheetXml);
            Assert.Contains("<xdr:twoCellAnchor editAs=\"twoCell\">", drawingXml);
            Assert.Contains("<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>", drawingXml);
            Assert.Contains("<xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>", drawingXml);
            Assert.Contains("<a:blip r:embed=\"rId1\"/>", drawingXml);
            Assert.Contains("Target=\"../media/image1.png\"", drawingRelsXml);
            Assert.Contains("PartName=\"/xl/drawings/drawing1.xml\"", contentTypesXml);
            Assert.Contains("Extension=\"png\" ContentType=\"image/png\"", contentTypesXml);

            var loaded = XlHelper.Load(path);
            Assert.Equal("header", loaded.Worksheets[0][0][0]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void AddImage_PathInput_UsesCellSpanForAnchor()
    {
        var imagePath = Path.Combine(Path.GetTempPath(), $"img-{Guid.NewGuid():N}.png");
        var xlsxPath = $"xlsx-img-path-{Guid.NewGuid():N}.xlsx";
        try
        {
            File.WriteAllBytes(imagePath, Png1x1Bytes);

            var wb = new XlWorkbook();
            var ws = wb.NewWorksheet("Sheet1");
            ws.AddImage(imagePath, rowIndex: 1, columnIndex: 2, rowSpan: 2, columnSpan: 3);

            XlHelper.Save(xlsxPath, wb);

            using var archive = ZipFile.OpenRead(xlsxPath);
            var drawingXml = ReadEntryText(archive, "xl/drawings/drawing1.xml");

            Assert.Contains("<xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>", drawingXml);
            Assert.Contains("<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>", drawingXml);
        }
        finally
        {
            if (File.Exists(imagePath)) File.Delete(imagePath);
            if (File.Exists(xlsxPath)) File.Delete(xlsxPath);
        }
    }

    [Fact]
    public void AddImage_UnsupportedExtension_ThrowsNotSupportedException()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");
        Assert.Throws<NotSupportedException>(() => ws.AddImage(Png1x1Bytes, ".webp", rowIndex: 0, columnIndex: 0));
    }

    [Fact]
    public void AddRealImagesTest()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");

        for (int i = 1; i <= 5; i++)
        {
            var imagePath = Path.Combine(AppContext.BaseDirectory, $"Images/{i}.jpg");
            ws.AddRow($"Image {i}");
            ws.AddImage(imagePath, rowIndex: i - 1, columnIndex: 1);
        }

        XlHelper.Save("test-images.xlsx", wb);
        Assert.True(File.Exists("test-images.xlsx"));
    }

    [Fact]
    public void AddRealImages_PlaceInCell_Test()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("Sheet1");

        var imagePath = Path.Combine(AppContext.BaseDirectory, $"Images/BASE_IMG_GAME.png");
        ws.AddRow($"Image 1");
        ws.AddImage(imagePath, rowIndex: 0, columnIndex: 1, placeInCell: true);

        XlHelper.Save("test-images-incell.xlsx", wb);
        Assert.True(File.Exists("test-images-incell.xlsx"));
    }

    private static string ReadEntryText(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path);
        Assert.NotNull(entry);
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
