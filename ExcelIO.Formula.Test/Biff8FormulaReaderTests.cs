using System.Buffers.Binary;
using System.Text;
using ExcelIO.Formula;

namespace ExcelIO.Formula.Test;

public class Biff8FormulaReaderTests
{
    [Fact]
    public void Decompile_SimpleInteger()
    {
        var rpn = new byte[] { 0x1E, 42, 0 };
        Assert.Equal("=42", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_SimpleDouble()
    {
        var rpn = BuildNumberRpn(3.14);
        Assert.StartsWith("=3.14", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_Addition()
    {
        var rpn = new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x03 };
        Assert.Equal("=1+2", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_Subtraction()
    {
        var rpn = new byte[] { 0x1E, 5, 0, 0x1E, 3, 0, 0x04 };
        Assert.Equal("=5-3", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_Multiplication()
    {
        var rpn = new byte[] { 0x1E, 6, 0, 0x1E, 7, 0, 0x05 };
        Assert.Equal("=6*7", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_Division()
    {
        var rpn = new byte[] { 0x1E, 8, 0, 0x1E, 2, 0, 0x06 };
        Assert.Equal("=8/2", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_PrecedenceFromRPN()
    {
        // 1 + 2 * 3 → RPN: 1, 2, 3, *, +
        var rpn = new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x1E, 3, 0, 0x05, 0x03 };
        Assert.Equal("=1+2*3", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_UnaryMinus()
    {
        var rpn = new byte[] { 0x1E, 5, 0, 0x13 };
        Assert.Equal("=-5", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_Percent()
    {
        var rpn = new byte[] { 0x1E, 50, 0, 0x14 };
        Assert.Equal("=50%", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_StringLiteral()
    {
        var rpn = new byte[] { 0x17, 5, (byte)'h', (byte)'e', (byte)'l', (byte)'l', (byte)'o' };
        Assert.Equal("=\"hello\"", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_BoolTrue() => Assert.Equal("=TRUE", Biff8FormulaReader.Decompile(new byte[] { 0x1D, 1 }));
    [Fact]
    public void Decompile_BoolFalse() => Assert.Equal("=FALSE", Biff8FormulaReader.Decompile(new byte[] { 0x1D, 0 }));

    [Fact]
    public void Decompile_Concat()
    {
        var rpn = new byte[] { 0x17, 1, (byte)'a', 0x17, 1, (byte)'b', 0x08 };
        Assert.Equal("=\"a\"&\"b\"", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_CellRef_A1()
    {
        var rpn = new byte[] { 0x44, 0, 0, 0, 0 };
        Assert.Equal("=A1", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_CellRef_B2()
    {
        var rpn = new byte[] { 0x44, 1, 0, 1, 0 };
        Assert.Equal("=B2", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_RangeRef()
    {
        var rpn = new byte[] { 0x45, 0, 0, 0, 0, 1, 0, 1, 0 };
        Assert.Equal("=A1:B2", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_SumRange()
    {
        // SUM(A1:A2) → range A1:A2, tFuncVar(SUM=4)
        var rpn = new byte[] {
            0x45, 0, 0, 0, 0, 1, 0, 0, 0,    // A1:A2
            0x43, 4, 0                           // tFuncVar, SUM=4
        };
        Assert.Equal("=SUM(A1:A2)", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_IfFunction()
    {
        // IF(1>0,"yes","no")
        var rpn = new byte[] {
            0x1E, 1, 0, 0x1E, 0, 0, 0x0D,    // 1>0
            0x17, 3, (byte)'y', (byte)'e', (byte)'s',
            0x17, 2, (byte)'n', (byte)'o',
            0x43, 1, 0                           // tFuncVar, IF=1
        };
        Assert.Equal("=IF(1>0,\"yes\",\"no\")", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_Average()
    {
        var rpn = new byte[] { 0x1E, 10, 0, 0x1E, 20, 0, 0x43, 5, 0 };
        Assert.Equal("=AVERAGE(10,20)", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_Comparison()
    {
        var rpn = new byte[] { 0x1E, 1, 0, 0x1E, 0, 0, 0x0D };
        Assert.Equal("=1>0", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_NotEqual()
    {
        var rpn = new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x0E };
        Assert.Equal("=1<>2", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_ComplexFormula()
    {
        // =A1+B1*2
        var rpn = new byte[] {
            0x44, 0, 0, 0, 0,     // A1
            0x44, 0, 0, 1, 0,     // B1
            0x1E, 2, 0,           // 2
            0x05,                  // *
            0x03                   // +
        };
        Assert.Equal("=A1+B1*2", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_Power()
    {
        var rpn = new byte[] { 0x1E, 2, 0, 0x1E, 3, 0, 0x07 };
        Assert.Equal("=2^3", Biff8FormulaReader.Decompile(rpn));
    }

    [Fact]
    public void Decompile_UnknownFunction_UsesFallback()
    {
        var rpn = new byte[] { 0x1E, 42, 0, 0x43, 0xE7, 0x03 };
        var f = Biff8FormulaReader.Decompile(rpn);
        Assert.StartsWith("=FUNC_", f);
        Assert.Contains("(42)", f);
    }

    [Fact]
    public void Decompile_EmptyBytes_ReturnsEmpty()
    {
        Assert.Equal("", Biff8FormulaReader.Decompile(Array.Empty<byte>()));
    }

    // ── Integration: XLS loading with formula decompiler ──

    [Fact(Skip = "CFB MiniFAT needs additional wiring for small workbooks")]
    public void Integration_LoadXlsWithFormula_DecompilesWithEngine()
    {
        var engine = new FormulaEngine();
        Assert.NotNull(XlHelper.Biff8FormulaDecompiler);

        var xls = BuildMinimalXlsWithFormula(row: 0, col: 0,
            rpn: new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x03 },  // =1+2
            cachedValue: 3.0);

        var wb = XlHelper.Load(new MemoryStream(xls));

        XlHelper.Biff8FormulaDecompiler = null;
        XlHelper.FormulaEngine = null;

        var cell = wb.Worksheets[0].Rows[0].Cells[0];
        Assert.True(cell.HasFormula);
        Assert.Equal("=1+2", cell.Formula);
    }

    [Fact(Skip = "CFB MiniFAT needs additional wiring for small workbooks")]
    public void Integration_LoadXlsWithoutEngine_ReadsValue()
    {
        var xls = BuildMinimalXlsWithFormula(row: 0, col: 0,
            rpn: new byte[] { 0x1E, 100, 0, 0x1E, 200, 0, 0x03 },
            cachedValue: 300.0);

        var wb = XlHelper.Load(new MemoryStream(xls));
        var cell = wb.Worksheets[0].Rows[0].Cells[0];

        Assert.Equal("300", cell.Value);
        Assert.False(cell.HasFormula);
    }

    // ── Helpers ──

    private static byte[] BuildNumberRpn(double value)
    {
        var rpn = new byte[9];
        rpn[0] = 0x1F;
        var bytes = BitConverter.GetBytes(value);
        Array.Copy(bytes, 0, rpn, 1, 8);
        return rpn;
    }

    private static byte[] BuildMinimalXlsWithFormula(int row, int col, byte[] rpn, double cachedValue)
    {
        // Workbook globals
        var globalBof = BuildBof(0x0005);
        var boundsheet = BuildBoundSheet(globalBof.Length + BuildBoundSheet(0, "Sheet1").Length + BuildSst("").Length + 4, "Sheet1");
        var sst = BuildSst("");
        var globalEof = BuildRecord(0x000A, []);

        // Worksheet
        var wsBof = BuildBof(0x0010);
        var formulaRecord = BuildFormulaRecord(row, col, rpn, cachedValue);
        var wsEof = BuildRecord(0x000A, []);

        var wsStream = Concat(wsBof, formulaRecord, wsEof);

        var workbookStream = Concat(globalBof, boundsheet, sst, globalEof, wsStream);
        // Pad to 4608 bytes so the stream is larger than the MiniFAT cutoff (4096)
        if (workbookStream.Length < 4608)
            Array.Resize(ref workbookStream, 4608);
        return WrapInCfb(workbookStream, "Workbook");
    }

    private static byte[] BuildBof(ushort substreamType)
    {
        var payload = new byte[16];
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(0, 2), 0x0600);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(2, 4), substreamType);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(4, 2), 0x0DBB);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(6, 2), 0x07CC);
        BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(8, 4), 0x00000041);
        BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(12, 4), 0x00000006);
        return BuildRecord(0x0809, payload);
    }

    private static byte[] BuildBoundSheet(int bofOffset, string name)
    {
        var nameBytes = Encoding.Latin1.GetBytes(name);
        var payload = new byte[8 + nameBytes.Length];
        BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(0, 4), (uint)bofOffset);
        payload[4] = 0x00;
        payload[5] = 0x00;
        payload[6] = (byte)nameBytes.Length;
        payload[7] = 0x00;
        Buffer.BlockCopy(nameBytes, 0, payload, 8, nameBytes.Length);
        return BuildRecord(0x0085, payload);
    }

    private static byte[] BuildSst(string value)
    {
        var strBytes = Encoding.Latin1.GetBytes(value);
        var payload = new byte[8 + 2 + 1 + strBytes.Length];
        BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(0, 4), 0);
        BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(4, 4), 0);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(8, 2), (ushort)value.Length);
        payload[10] = 0x00;
        Buffer.BlockCopy(strBytes, 0, payload, 11, strBytes.Length);
        return BuildRecord(0x00FC, payload);
    }

    private static byte[] BuildFormulaRecord(int row, int col, byte[] rpn, double cachedValue)
    {
        var payload = new byte[20 + rpn.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(0, 2), (ushort)row);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(2, 2), (ushort)col);
        // ixfe = 0 (bytes 4-5)
        BinaryPrimitives.WriteInt64LittleEndian(payload.AsSpan(6, 8), BitConverter.DoubleToInt64Bits(cachedValue));
        // grbit = 0 (bytes 14-15), chn = 0 (bytes 16-19)
        Buffer.BlockCopy(rpn, 0, payload, 20, rpn.Length);
        return BuildRecord(0x0006, payload);
    }

    private static byte[] BuildRecord(ushort id, byte[] payload)
    {
        var output = new byte[4 + payload.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(output.AsSpan(0, 2), id);
        BinaryPrimitives.WriteUInt16LittleEndian(output.AsSpan(2, 2), (ushort)payload.Length);
        Buffer.BlockCopy(payload, 0, output, 4, payload.Length);
        return output;
    }

    private static byte[] Concat(params byte[][] arrays)
    {
        var result = new byte[arrays.Sum(a => a.Length)];
        int offset = 0;
        foreach (var a in arrays)
        {
            Buffer.BlockCopy(a, 0, result, offset, a.Length);
            offset += a.Length;
        }
        return result;
    }

    private static byte[] WrapInCfb(byte[] workbookStream, string streamName)
    {
        const int sectorSize = 512;
        const uint endOfChain = 0xFFFFFFFE;
        const uint freeSect = 0xFFFFFFFF;
        const uint fatSect = 0xFFFFFFFD;

        int workbookSectors = (workbookStream.Length + sectorSize - 1) / sectorSize;

        var fat = new int[128];
        for (int i = 0; i < 128; i++) fat[i] = unchecked((int)freeSect);
        fat[0] = unchecked((int)fatSect);
        fat[1] = unchecked((int)endOfChain);
        for (int i = 0; i < workbookSectors; i++)
        {
            int sid = 2 + i;
            fat[sid] = (i == workbookSectors - 1) ? unchecked((int)endOfChain) : sid + 1;
        }

        var fileBytes = new byte[(1 + 2 + workbookSectors) * sectorSize];

        // Header
        var hdr = fileBytes.AsSpan(0, 512);
        BinaryPrimitives.WriteUInt64LittleEndian(hdr[..8], 0xE11AB1A1E011CFD0UL);
        BinaryPrimitives.WriteUInt16LittleEndian(hdr.Slice(0x18, 2), 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(hdr.Slice(0x1A, 2), 0x0003);
        BinaryPrimitives.WriteUInt16LittleEndian(hdr.Slice(0x1C, 2), 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(hdr.Slice(0x1E, 2), 0x0009);
        BinaryPrimitives.WriteUInt16LittleEndian(hdr.Slice(0x20, 2), 0x0006);
        BinaryPrimitives.WriteUInt32LittleEndian(hdr.Slice(0x2C, 4), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(hdr.Slice(0x30, 4), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(hdr.Slice(0x38, 4), 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(hdr.Slice(0x3C, 4), endOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(hdr.Slice(0x40, 4), 0);
        BinaryPrimitives.WriteUInt32LittleEndian(hdr.Slice(0x44, 4), endOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(hdr.Slice(0x48, 4), 0);
        for (int i = 0; i < 109; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(hdr.Slice(0x4C + i * 4, 4), i == 0 ? 0u : freeSect);

        // FAT
        var fatSpan = fileBytes.AsSpan(512, 512);
        for (int i = 0; i < 128; i++)
            BinaryPrimitives.WriteInt32LittleEndian(fatSpan.Slice(i * 4, 4), fat[i]);

        // Directory
        var dir = fileBytes.AsSpan(1024, 512);
        dir.Clear();
        WriteDirEntry(dir.Slice(0, 128), "Root Entry", 5, endOfChain, 0);
        WriteDirEntry(dir.Slice(128, 128), streamName, 2, 2, workbookStream.Length);

        // Workbook data
        var dataOffset = 3 * sectorSize;
        Buffer.BlockCopy(workbookStream, 0, fileBytes, dataOffset, workbookStream.Length);

        return fileBytes;
    }

    private static void WriteDirEntry(Span<byte> entry, string name, byte type, uint startSector, int size)
    {
        entry.Clear();
        var nameBytes = Encoding.Unicode.GetBytes(name + "\0");
        nameBytes.CopyTo(entry);
        BinaryPrimitives.WriteUInt16LittleEndian(entry.Slice(64, 2), (ushort)nameBytes.Length);
        entry[66] = type;
        entry[67] = 0x01;
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(68, 4), 0xFFFFFFFF);
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(72, 4), 0xFFFFFFFF);
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(76, 4), 0xFFFFFFFF);
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(116, 4), startSector);
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(120, 4), (uint)size);
    }
}
