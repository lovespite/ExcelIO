namespace ExcelIO.Test;

public class XlRangeTests
{
    [Fact]
    public void Range_A1_Parsing()
    {
        var ws = new XlWorksheet(new XlWorkbook());
        var range = ws.Range("A1");
        Assert.Equal(0, range.StartColumn);
        Assert.Equal(0, range.StartRow);
        Assert.Equal(0, range.EndColumn);
        Assert.Equal(0, range.EndRow);
        Assert.False(range.IsInfiniteRow);
        Assert.False(range.IsInfiniteColumn);
        Assert.Equal(new XlPoint(0, 0), range.StartPosition);
        Assert.Equal(new XlPoint(0, 0), range.EndPosition);
    }

    [Fact]
    public void Range_A_Parsing()
    {
        var ws = new XlWorksheet(new XlWorkbook());
        var range = ws.Range("A");
        Assert.Equal(0, range.StartColumn);
        Assert.Equal(0, range.EndColumn);
        Assert.True(range.IsInfiniteRow);
        Assert.Equal(new XlPoint(0, 0), range.StartPosition);
        Assert.Equal(new XlPoint(-1, -1), range.EndPosition);
    }

    [Fact]
    public void Range_AtoC_Parsing()
    {
        var ws = new XlWorksheet(new XlWorkbook());
        var range = ws.Range("A:C");
        Assert.Equal(0, range.StartColumn);
        Assert.Equal(2, range.EndColumn);
        Assert.True(range.IsInfiniteRow);
        Assert.Equal(new XlPoint(0, 0), range.StartPosition);
        Assert.Equal(new XlPoint(-1, -1), range.EndPosition);
    }

    [Fact]
    public void Range_1to3_Parsing()
    {
        var ws = new XlWorksheet(new XlWorkbook());
        var range = ws.Range("1:3");
        Assert.Equal(0, range.StartRow);
        Assert.Equal(2, range.EndRow);
        Assert.True(range.IsInfiniteColumn);
        Assert.Equal(new XlPoint(0, 0), range.StartPosition);
        Assert.Equal(new XlPoint(-1, -1), range.EndPosition);
    }

    [Fact]
    public void Range_A1toC3_Parsing()
    {
        var ws = new XlWorksheet(new XlWorkbook());
        var range = ws.Range("A1:C3");
        Assert.Equal(0, range.StartColumn);
        Assert.Equal(0, range.StartRow);
        Assert.Equal(2, range.EndColumn);
        Assert.Equal(2, range.EndRow);
        Assert.False(range.IsInfiniteRow);
        Assert.False(range.IsInfiniteColumn);
        Assert.Equal(new XlPoint(0, 0), range.StartPosition);
        Assert.Equal(new XlPoint(2, 2), range.EndPosition);
    }

    [Fact]
    public void Range_SetContent_MemorySafe()
    {
        var ws = new XlWorksheet(new XlWorkbook());
        ws.Range("A").SetContent("Hello");
        Assert.Empty(ws.Rows);
        
        ws.Range("A1:B2").SetContent("World");
        Assert.Equal(2, ws.Rows.Count);
        Assert.Equal("World", ws.Rows[0].Cells[0].Value);
        Assert.Equal("World", ws.Rows[1].Cells[1].Value);
    }

    [Fact]
    public void Range_Merge_Unmerge()
    {
        var ws = new XlWorksheet(new XlWorkbook());
        var range = ws.Range("A1:C3");
        range.Merge();
        Assert.Contains("A1:C3", ws.MergedCells);
        
        range.Unmerge();
        Assert.DoesNotContain("A1:C3", ws.MergedCells);
    }

    [Fact]
    public void Range_AddImage_CalculatesCorrectSpan()
    {
        var ws = new XlWorksheet(new XlWorkbook());
        var range = ws.Range("A1:C3");
        // Using a dummy byte array for test
        var img = ws.AddImage(new byte[10], "jpg", range);
        
        Assert.Equal(0, img.RowIndex);
        Assert.Equal(0, img.ColumnIndex);
        Assert.Equal(3, img.RowSpan);
        Assert.Equal(3, img.ColumnSpan);
    }

    [Fact]
    public void Range_Expression_Property()
    {
        var ws = new XlWorksheet(new XlWorkbook());
        Assert.Equal("A1", ws.Range("A1").RangeExpression);
        Assert.Equal("A:C", ws.Range("A:C").RangeExpression);
        Assert.Equal("1:3", ws.Range("1:3").RangeExpression);
    }
}
