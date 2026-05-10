namespace ExcelIO.Test;

public class XlStyleTests
{
    [Fact]
    public void TestXlsxStyling()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("StyleTest");

        // 1. Sheet Options
        ws.Options.TabColor = "FFFF0000"; // Red tab
        ws.Options.ShowGridLines = false;
        ws.Options.DefaultRowHeight = 30;

        // 2. Column Widths and Styles
        ws.Columns[0] = new XlColumn { Width = 20 }; // Col A width 20
        ws.Columns[1] = new XlColumn { Width = 30, Style = new XlStyle { FillColor = "FF00FF00" } }; // Col B width 30, green fill
        ws.Columns[2] = new XlColumn { Hidden = true }; // Col C hidden

        // 3. Row Heights and Styles
        var row1 = ws.NewRow(["Header 1", "Header 2", "Header 3"]);
        row1.Height = 40;
        row1.Style = new XlStyle 
        { 
            Bold = true, 
            FontSize = 14, 
            FontColor = "FFFFFFFF", 
            FillColor = "FF0000FF",
            Alignment = new XlAlignment { Horizontal = XlHorizontalAlignment.Center, Vertical = XlVerticalAlignment.Center }
        };

        var row2 = ws.NewRow(["Cell 2.1", "Cell 2.2", "Cell 2.3"]);
        row2[0] = "Override Row Style";
        ws.Rows[1].Cells[0].Style = new XlStyle { Bold = false, FontColor = "FFFF0000" }; // Red text, not bold

        var row3 = ws.NewRow(["Hidden Row", "Hidden", "Row"]);
        row3.Hidden = true;

        var row4 = ws.NewRow(["Borders", "Test", ""]);
        row4.Cells[0].Style = new XlStyle
        {
            Border = new XlBorder
            {
                Bottom = XlBorderStyle.Thick,
                BottomColor = "FFFF0000",
                Right = XlBorderStyle.Double,
                RightColor = "FF00FF00"
            }
        };

        var outputPath = "StyleTest.xlsx";
        XlHelper.Save(outputPath, wb);

        Assert.True(File.Exists(outputPath));

        // Round-trip verification
        var loadedWb = XlHelper.Load(outputPath);
        var loadedWs = loadedWb.Worksheets[0];

        Assert.Equal("StyleTest", loadedWs.Name);
        Assert.Equal("FFFF0000", loadedWs.Options.TabColor);
        Assert.False(loadedWs.Options.ShowGridLines);
        Assert.Equal(30, loadedWs.Options.DefaultRowHeight);

        Assert.Equal(20, loadedWs.Columns[0].Width);
        Assert.Equal(30, loadedWs.Columns[1].Width);
        Assert.Equal("FF00FF00", loadedWs.Columns[1].Style?.FillColor);
        Assert.True(loadedWs.Columns[2].Hidden);

        Assert.Equal(40, loadedWs.Rows[0].Height);
        Assert.True(loadedWs.Rows[0].Style?.Bold);
        Assert.Equal(14, loadedWs.Rows[0].Style?.FontSize);
        Assert.Equal("FFFFFFFF", loadedWs.Rows[0].Style?.FontColor);
        Assert.Equal("FF0000FF", loadedWs.Rows[0].Style?.FillColor);
        Assert.Equal(XlHorizontalAlignment.Center, loadedWs.Rows[0].Style?.Alignment?.Horizontal);

        Assert.Equal("Override Row Style", loadedWs.Rows[1].Cells[0].Value);
        Assert.False(loadedWs.Rows[1].Cells[0].Style?.Bold);
        Assert.Equal("FFFF0000", loadedWs.Rows[1].Cells[0].Style?.FontColor);

        Assert.True(loadedWs.Rows[2].Hidden);

        Assert.Equal(XlBorderStyle.Thick, loadedWs.Rows[3].Cells[0].Style?.Border?.Bottom);
        Assert.Equal("FFFF0000", loadedWs.Rows[3].Cells[0].Style?.Border?.BottomColor);
    }

    [Fact]
    public void TestXlsxImageRoundTrip()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("ImageTest");
        
        var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "1.jpg");
        if (!File.Exists(imagePath)) imagePath = @"D:\projects\ExcelIO\ExcelIO.Test\Images\1.jpg";
        
        ws.AddImage(imagePath, 2, 2, 5, 3);
        
        var outputPath = "ImageRoundTrip.xlsx";
        XlHelper.Save(outputPath, wb);
        
        var loadedWb = XlHelper.Load(outputPath);
        var loadedWs = loadedWb.Worksheets[0];
        
        Assert.Single(loadedWs.Images);
        var img = loadedWs.Images[0];
        Assert.Equal(2, img.RowIndex);
        Assert.Equal(2, img.ColumnIndex);
        Assert.Equal(5, img.RowSpan);
        Assert.Equal(3, img.ColumnSpan);
        Assert.Equal("jpg", img.Extension);
        Assert.True(img.Bytes.Length > 0);
    }

    [Fact]
    public void TestXlsxLoadOptions()
    {
        var wb = new XlWorkbook();
        var ws = wb.NewWorksheet("OptionTest");
        ws.Columns[0] = new XlColumn { Width = 25, Style = new XlStyle { Bold = true } };
        var row = ws.NewRow(["Styled"]);
        row.Style = new XlStyle { FillColor = "FFFF0000" };
        
        var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "1.jpg");
        if (!File.Exists(imagePath)) imagePath = @"D:\projects\ExcelIO\ExcelIO.Test\Images\1.jpg";
        ws.AddImage(imagePath, 2, 2, 1, 1);
        
        var outputPath = "LoadOptionsTest.xlsx";
        XlHelper.Save(outputPath, wb);
        
        // 1. Load without styles
        var noStylesWb = XlHelper.Load(outputPath, new XlLoadOptions { LoadStyles = false, LoadImages = true });
        var noStylesWs = noStylesWb.Worksheets[0];
        Assert.Null(noStylesWs.Columns[0].Style);
        Assert.Null(noStylesWs.Rows[0].Style);
        Assert.Single(noStylesWs.Images);
        
        // 2. Load without images
        var noImagesWb = XlHelper.Load(outputPath, new XlLoadOptions { LoadStyles = true, LoadImages = false });
        var noImagesWs = noImagesWb.Worksheets[0];
        Assert.NotNull(noImagesWs.Columns[0].Style);
        Assert.Empty(noImagesWs.Images);
    }
}
