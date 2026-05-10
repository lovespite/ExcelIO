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
    }
}
