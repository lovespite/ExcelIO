
using System.Buffers.Binary;
using System.Text;

namespace ExcelIO.Tests;

public class XlHelperXlsLoadTests
{
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSect = 0xFFFFFFFF;
    private const uint FatSect = 0xFFFFFFFD;

    [Fact]
    public void Load_Xls_ReadsMinimalTextAndNumberCells()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xls-min-{Guid.NewGuid():N}.xls");
        try
        {
            File.WriteAllBytes(path, BuildMinimalXls());

            var wb = XlHelper.Load(path);
            Assert.Single(wb.Worksheets);
            Assert.Equal("Sheet1", wb.Worksheets[0].Name);
            Assert.Equal("123.5", wb.Worksheets[0][0][0]);
            Assert.Equal("hello", wb.Worksheets[0][1][1]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Load_Stream_Xls_ReadsMinimalTextAndNumberCells()
    {
        var bytes = BuildMinimalXls();
        using var stream = new MemoryStream(bytes, writable: false);
        var wb = XlHelper.Load(stream);
        Assert.Single(wb.Worksheets);
        Assert.Equal("Sheet1", wb.Worksheets[0].Name);
        Assert.Equal("123.5", wb.Worksheets[0][0][0]);
        Assert.Equal("hello", wb.Worksheets[0][1][1]);
    }

    [Fact]
    public void Load_Stream_Xlsx_StillWorks()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-stream-{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = new XlWorkbook();
            wb.NewWorksheet("SheetA").AddRow("A", "B");
            XlHelper.Save(path, wb);

            var bytes = File.ReadAllBytes(path);
            using var stream = new MemoryStream(bytes, writable: false);
            var loaded = XlHelper.Load(stream);
            Assert.Single(loaded.Worksheets);
            Assert.Equal("SheetA", loaded.Worksheets[0].Name);
            Assert.Equal("A", loaded.Worksheets[0][0][0]);
            Assert.Equal("B", loaded.Worksheets[0][0][1]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Load_Stream_UnknownHeader_ThrowsNotSupportedException()
    {
        using var stream = new MemoryStream("bad-stream"u8.ToArray(), writable: false);
        Assert.Throws<NotSupportedException>(() => XlHelper.Load(stream));
    }

    [Fact]
    public void Load_Stream_NonSeekableXls_Works()
    {
        var bytes = BuildMinimalXls();
        using var stream = new NonSeekableReadStream(bytes);
        var wb = XlHelper.Load(stream);
        Assert.Equal("123.5", wb.Worksheets[0][0][0]);
        Assert.Equal("hello", wb.Worksheets[0][1][1]);
    }

    [Fact]
    public async Task LoadAsync_Stream_Xls_ReadsMinimalTextAndNumberCells()
    {
        var bytes = BuildMinimalXls();
        await using var stream = new MemoryStream(bytes, writable: false);
        var wb = await XlHelper.LoadAsync(stream, ".xls");
        Assert.Equal("123.5", wb.Worksheets[0][0][0]);
        Assert.Equal("hello", wb.Worksheets[0][1][1]);
    }

    [Fact]
    public void Load_InvalidXls_ThrowsFormatException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xls-bad-{Guid.NewGuid():N}.xls");
        try
        {
            File.WriteAllBytes(path, "not-an-xls"u8.ToArray());
            Assert.Throws<FormatException>(() => XlHelper.Load(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Load_Xlsx_StillWorks()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = new XlWorkbook();
            wb.NewWorksheet("SheetA").AddRow("A", "B");
            XlHelper.Save(path, wb);

            var loaded = XlHelper.Load(path);
            Assert.Single(loaded.Worksheets);
            Assert.Equal("SheetA", loaded.Worksheets[0].Name);
            Assert.Equal("A", loaded.Worksheets[0][0][0]);
            Assert.Equal("B", loaded.Worksheets[0][0][1]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static byte[] BuildMinimalXls()
    {
        var globalBof = BuildBofRecord(0x0005);
        var sst = BuildSstRecord("hello");
        var eof = BuildRecord(0x000A, []);

        var worksheet = BuildWorksheetStream();
        var boundsheet = BuildBoundSheetRecord(globalBof.Length + BuildBoundSheetRecord(0, "Sheet1").Length + sst.Length + eof.Length, "Sheet1");

        var workbookStream = new byte[globalBof.Length + boundsheet.Length + sst.Length + eof.Length + worksheet.Length];
        var cursor = 0;
        Buffer.BlockCopy(globalBof, 0, workbookStream, cursor, globalBof.Length); cursor += globalBof.Length;
        Buffer.BlockCopy(boundsheet, 0, workbookStream, cursor, boundsheet.Length); cursor += boundsheet.Length;
        Buffer.BlockCopy(sst, 0, workbookStream, cursor, sst.Length); cursor += sst.Length;
        Buffer.BlockCopy(eof, 0, workbookStream, cursor, eof.Length); cursor += eof.Length;
        Buffer.BlockCopy(worksheet, 0, workbookStream, cursor, worksheet.Length);

        if (workbookStream.Length < 4608)
        {
            Array.Resize(ref workbookStream, 4608);
        }
        return WrapInCfb(workbookStream, "Workbook");
    }

    private static byte[] BuildWorksheetStream()
    {
        var bof = BuildBofRecord(0x0010);
        var number = BuildNumberRecord(row: 0, col: 0, value: 123.5d);
        var labelSst = BuildLabelSstRecord(row: 1, col: 1, sstIndex: 0);
        var eof = BuildRecord(0x000A, []);

        var stream = new byte[bof.Length + number.Length + labelSst.Length + eof.Length];
        var cursor = 0;
        Buffer.BlockCopy(bof, 0, stream, cursor, bof.Length); cursor += bof.Length;
        Buffer.BlockCopy(number, 0, stream, cursor, number.Length); cursor += number.Length;
        Buffer.BlockCopy(labelSst, 0, stream, cursor, labelSst.Length); cursor += labelSst.Length;
        Buffer.BlockCopy(eof, 0, stream, cursor, eof.Length);
        return stream;
    }

    private static byte[] BuildBofRecord(ushort substreamType)
    {
        Span<byte> payload = stackalloc byte[16];
        BinaryPrimitives.WriteUInt16LittleEndian(payload[0..2], 0x0600);
        BinaryPrimitives.WriteUInt16LittleEndian(payload[2..4], substreamType);
        BinaryPrimitives.WriteUInt16LittleEndian(payload[4..6], 0x0DBB);
        BinaryPrimitives.WriteUInt16LittleEndian(payload[6..8], 0x07CC);
        BinaryPrimitives.WriteUInt32LittleEndian(payload[8..12], 0x00000041);
        BinaryPrimitives.WriteUInt32LittleEndian(payload[12..16], 0x00000006);
        return BuildRecord(0x0809, payload.ToArray());
    }

    private static byte[] BuildBoundSheetRecord(int bofOffset, string sheetName)
    {
        var nameBytes = Encoding.Latin1.GetBytes(sheetName);
        var payload = new byte[8 + nameBytes.Length];
        BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(0, 4), (uint)bofOffset);
        payload[4] = 0x00;
        payload[5] = 0x00;
        payload[6] = (byte)nameBytes.Length;
        payload[7] = 0x00;
        Buffer.BlockCopy(nameBytes, 0, payload, 8, nameBytes.Length);
        return BuildRecord(0x0085, payload);
    }

    private static byte[] BuildSstRecord(string value)
    {
        var strBytes = Encoding.Latin1.GetBytes(value);
        var payload = new byte[8 + 2 + 1 + strBytes.Length];
        BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(0, 4), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(4, 4), 1);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(8, 2), (ushort)value.Length);
        payload[10] = 0x00;
        Buffer.BlockCopy(strBytes, 0, payload, 11, strBytes.Length);
        return BuildRecord(0x00FC, payload);
    }

    private static byte[] BuildNumberRecord(ushort row, ushort col, double value)
    {
        var payload = new byte[14];
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(0, 2), row);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(2, 2), col);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(4, 2), 0);
        BinaryPrimitives.WriteInt64LittleEndian(payload.AsSpan(6, 8), BitConverter.DoubleToInt64Bits(value));
        return BuildRecord(0x0203, payload);
    }

    private static byte[] BuildLabelSstRecord(ushort row, ushort col, uint sstIndex)
    {
        var payload = new byte[10];
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(0, 2), row);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(2, 2), col);
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(4, 2), 0);
        BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(6, 4), sstIndex);
        return BuildRecord(0x00FD, payload);
    }

    private static byte[] BuildRecord(ushort id, byte[] payload)
    {
        var output = new byte[4 + payload.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(output.AsSpan(0, 2), id);
        BinaryPrimitives.WriteUInt16LittleEndian(output.AsSpan(2, 2), (ushort)payload.Length);
        Buffer.BlockCopy(payload, 0, output, 4, payload.Length);
        return output;
    }

    private static byte[] WrapInCfb(byte[] workbookStream, string workbookStreamName)
    {
        const int sectorSize = 512;
        var workbookSectorCount = (workbookStream.Length + sectorSize - 1) / sectorSize;
        var totalSectors = 2 + workbookSectorCount; // FAT + DIR + Workbook

        var fat = new int[128];
        for (int i = 0; i < fat.Length; i++) fat[i] = unchecked((int)FreeSect);
        fat[0] = unchecked((int)FatSect);
        fat[1] = unchecked((int)EndOfChain);
        for (int i = 0; i < workbookSectorCount; i++)
        {
            var sid = 2 + i;
            fat[sid] = i == workbookSectorCount - 1 ? unchecked((int)EndOfChain) : sid + 1;
        }

        var fileBytes = new byte[(1 + totalSectors) * sectorSize];
        WriteHeader(fileBytes);
        WriteFatSector(fileBytes, fat);
        WriteDirectorySector(fileBytes, workbookStream.Length, workbookStreamName);

        var workbookOffset = (2 + 1) * sectorSize;
        Buffer.BlockCopy(workbookStream, 0, fileBytes, workbookOffset, workbookStream.Length);
        return fileBytes;
    }

    private static void WriteHeader(byte[] fileBytes)
    {
        var header = fileBytes.AsSpan(0, 512);
        BinaryPrimitives.WriteUInt64LittleEndian(header[..8], 0xE11AB1A1E011CFD0UL);
        BinaryPrimitives.WriteUInt16LittleEndian(header.Slice(0x18, 2), 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(header.Slice(0x1A, 2), 0x0003);
        BinaryPrimitives.WriteUInt16LittleEndian(header.Slice(0x1C, 2), 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(header.Slice(0x1E, 2), 0x0009);
        BinaryPrimitives.WriteUInt16LittleEndian(header.Slice(0x20, 2), 0x0006);
        BinaryPrimitives.WriteUInt32LittleEndian(header.Slice(0x2C, 4), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header.Slice(0x30, 4), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header.Slice(0x38, 4), 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(header.Slice(0x3C, 4), EndOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(header.Slice(0x40, 4), 0);
        BinaryPrimitives.WriteUInt32LittleEndian(header.Slice(0x44, 4), EndOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(header.Slice(0x48, 4), 0);
        for (int i = 0; i < 109; i++)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(header.Slice(0x4C + i * 4, 4), i == 0 ? 0u : FreeSect);
        }
    }

    private static void WriteFatSector(byte[] fileBytes, int[] fat)
    {
        var fatOffset = (0 + 1) * 512;
        var fatBytes = fileBytes.AsSpan(fatOffset, 512);
        for (int i = 0; i < 128; i++)
        {
            BinaryPrimitives.WriteInt32LittleEndian(fatBytes.Slice(i * 4, 4), fat[i]);
        }
    }

    private static void WriteDirectorySector(byte[] fileBytes, int workbookSize, string workbookName)
    {
        var dirOffset = (1 + 1) * 512;
        var dirBytes = fileBytes.AsSpan(dirOffset, 512);
        dirBytes.Clear();

        WriteDirectoryEntry(dirBytes.Slice(0, 128), "Root Entry", type: 5, startSector: EndOfChain, size: 0);
        WriteDirectoryEntry(dirBytes.Slice(128, 128), workbookName, type: 2, startSector: 2, size: workbookSize);
    }

    private static void WriteDirectoryEntry(Span<byte> entry, string name, byte type, uint startSector, int size)
    {
        entry.Clear();

        var nameWithNull = name + "\0";
        var nameBytes = Encoding.Unicode.GetBytes(nameWithNull);
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

    private sealed class NonSeekableReadStream : Stream
    {
        private readonly MemoryStream _inner;

        public NonSeekableReadStream(byte[] bytes)
        {
            _inner = new MemoryStream(bytes, writable: false);
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override int Read(Span<byte> buffer) => _inner.Read(buffer);
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) =>
            _inner.ReadAsync(buffer, cancellationToken);
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            _inner.ReadAsync(buffer, offset, count, cancellationToken);
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }

        public override async ValueTask DisposeAsync()
        {
            await _inner.DisposeAsync();
            await base.DisposeAsync();
        }
    }
}
