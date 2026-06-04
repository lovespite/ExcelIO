using System.Buffers.Binary;
using System.Text;

namespace ExcelIO.NetStandard.Test;

public class SpanExtensionsTests
{
    [Fact]
    public void Encoding_GetString_ReadOnlySpan_ReturnsCorrectString()
    {
        var bytes = "Hello World"u8;
        var result = Encoding.UTF8.GetString(bytes);
        Assert.Equal("Hello World", result);
    }

    [Fact]
    public void Encoding_GetString_EmptySpan_ReturnsEmpty()
    {
        var result = Encoding.UTF8.GetString([]);
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void Encoding_GetString_Unicode_ReturnsCorrectString()
    {
        var str = "你好世界";
        var bytes = Encoding.Unicode.GetBytes(str);
        var result = Encoding.Unicode.GetString(bytes.AsSpan());
        Assert.Equal(str, result);
    }

    [Fact]
    public void Encoding_GetString_Latin1_ReturnsCorrectString()
    {
        var str = "café";
        var bytes = Encoding.GetEncoding(28591).GetBytes(str);
        var result = Encoding.GetEncoding(28591).GetString(bytes.AsSpan());
        Assert.Equal(str, result);
    }

    [Fact]
    public void MemoryStream_Write_ReadOnlySpan_WritesCorrectly()
    {
        using var ms = new MemoryStream();
        var data = "test data"u8;
        ms.Write(data);
        Assert.Equal(data.Length, ms.Length);
        Assert.Equal("test data", Encoding.UTF8.GetString(ms.ToArray()));
    }

    [Fact]
    public void MemoryStream_Write_MultipleSpans_AppendsCorrectly()
    {
        using var ms = new MemoryStream();
        var part1 = "hello"u8;
        var part2 = " world"u8;
        ms.Write(part1);
        ms.Write(part2);
        Assert.Equal("hello world", Encoding.UTF8.GetString(ms.ToArray()));
    }

    [Fact]
    public void MemoryStream_Write_EmptySpan_NoOp()
    {
        using var ms = new MemoryStream();
        ms.Write(ReadOnlySpan<byte>.Empty);
        Assert.Equal(0, ms.Length);
    }

    [Fact]
    public void MemoryStream_Write_AfterSeek_WritesAtPosition()
    {
        using var ms = new MemoryStream();
        var prefix = "prefix"u8;
        ms.Write(prefix);
        ms.Seek(0, SeekOrigin.Begin);
        var overwrite = "overwr"u8;
        ms.Write(overwrite);
        Assert.Equal("overwr", Encoding.UTF8.GetString(ms.ToArray()));
    }

    [Fact]
    public void MemoryStream_Write_LargeSpan_WritesCorrectly()
    {
        using var ms = new MemoryStream();
        var data = new byte[10000];
        new Random(42).NextBytes(data);
        ms.Write(data.AsSpan());
        var result = ms.ToArray();
        Assert.Equal(data, result);
    }
}
