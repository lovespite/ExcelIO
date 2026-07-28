using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace ExcelIO;

internal static class XlsBiff8Reader
{
    private const ushort RecordBof = 0x0809;
    private const ushort RecordEof = 0x000A;
    private const ushort RecordBoundSheet = 0x0085;
    private const ushort RecordSst = 0x00FC;
    private const ushort RecordContinue = 0x003C;
    private const ushort RecordLabelSst = 0x00FD;
    private const ushort RecordNumber = 0x0203;
    private const ushort RecordRk = 0x027E;
    private const ushort RecordLabel = 0x0204;
    private const ushort RecordFormula = 0x0006;
    private const ushort RecordShrFmla = 0x04BC;

    public static XlWorkbook Load(ReadOnlySpan<byte> workbookStream)
    {
        if (workbookStream.Length < 4)
        {
            throw new FormatException("Invalid XLS file: workbook stream is too short.");
        }

        var sheets = new List<BoundSheet>();
        var sst = new List<string>();
        int offset = 0;
        while (TryReadRecord(workbookStream, offset, out var id, out var payloadOffset, out var payloadLength, out var next))
        {
            if (id == RecordBoundSheet)
            {
                sheets.Add(ParseBoundSheet(workbookStream.Slice(payloadOffset, payloadLength)));
            }
            else if (id == RecordSst)
            {
                var merged = MergeContinuePayload(workbookStream, payloadOffset, payloadLength, next, out var nextAfterContinue);
                sst = ParseSst(merged);
                next = nextAfterContinue;
            }

            offset = next;
            if (id == RecordEof && sheets.Count > 0)
            {
                break;
            }
        }

        if (sheets.Count == 0)
        {
            throw new FormatException("Invalid XLS file: no worksheet metadata found.");
        }

        var wb = new XlWorkbook();
        for (int i = 0; i < sheets.Count; i++)
        {
            var sheet = sheets[i];
            var ws = wb.NewWorksheet(string.IsNullOrWhiteSpace(sheet.Name) ? $"Sheet{i + 1}" : sheet.Name);
            ParseWorksheet(workbookStream, sheet.BofOffset, ws, sst);
        }
        return wb;
    }

    private static void ParseWorksheet(ReadOnlySpan<byte> stream, int sheetOffset, XlWorksheet ws, IReadOnlyList<string> sst)
    {
        if (!TryReadRecord(stream, sheetOffset, out var id, out _, out _, out var next) || id != RecordBof)
        {
            throw new FormatException("Invalid XLS file: worksheet BOF not found.");
        }

        // Shared formula tracking: (firstRow, firstCol, lastRow, lastCol) -> RPN bytes
        var sharedFormulas = new Dictionary<(int, int, int, int), byte[]>();

        int offset = next;
        while (TryReadRecord(stream, offset, out id, out var payloadOffset, out var payloadLength, out next))
        {
            var payload = stream.Slice(payloadOffset, payloadLength);
            switch (id)
            {
                case RecordNumber:
                    if (payload.Length >= 14)
                    {
                        var row = ReadUInt16(payload, 0);
                        var col = ReadUInt16(payload, 2);
                        var value = BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(payload.Slice(6, 8)));
                        SetCell(ws, row, col, value.ToString("G17", CultureInfo.InvariantCulture));
                    }
                    break;
                case RecordRk:
                    if (payload.Length >= 10)
                    {
                        var row = ReadUInt16(payload, 0);
                        var col = ReadUInt16(payload, 2);
                        var rk = BinaryPrimitives.ReadUInt32LittleEndian(payload.Slice(6, 4));
                        var value = DecodeRk(rk).ToString("G17", CultureInfo.InvariantCulture);
                        SetCell(ws, row, col, value);
                    }
                    break;
                case RecordLabelSst:
                    if (payload.Length >= 10)
                    {
                        var row = ReadUInt16(payload, 0);
                        var col = ReadUInt16(payload, 2);
                        var idx = (int)BinaryPrimitives.ReadUInt32LittleEndian(payload.Slice(6, 4));
                        var value = idx >= 0 && idx < sst.Count ? sst[idx] : string.Empty;
                        SetCell(ws, row, col, value);
                    }
                    break;
                case RecordLabel:
                    if (payload.Length >= 8)
                    {
                        var row = ReadUInt16(payload, 0);
                        var col = ReadUInt16(payload, 2);
                        var len = ReadUInt16(payload, 6);
                        var available = Math.Min((int)len, Math.Max(0, payload.Length - 8));
                        var value = Encoding.GetEncoding(28591).GetString(payload.Slice(8, available));
                        SetCell(ws, row, col, value);
                    }
                    break;
                case RecordFormula:
                    ParseFormulaRecord(ws, payload, sharedFormulas);
                    break;
                case RecordShrFmla:
                    ParseSharedFormulaRecord(payload, sharedFormulas);
                    break;
                case RecordEof:
                    return;
            }

            offset = next;
        }
    }

    private static byte[] MergeContinuePayload(
        ReadOnlySpan<byte> stream,
        int firstPayloadOffset,
        int firstPayloadLength,
        int nextRecordOffset,
        out int nextAfterContinue)
    {
        using var ms = new MemoryStream(firstPayloadLength + 256);
        ms.Write(stream.Slice(firstPayloadOffset, firstPayloadLength));

        int offset = nextRecordOffset;
        while (TryReadRecord(stream, offset, out var id, out var payloadOffset, out var payloadLength, out var next))
        {
            if (id != RecordContinue)
            {
                nextAfterContinue = offset;
                return ms.ToArray();
            }
            ms.Write(stream.Slice(payloadOffset, payloadLength));
            offset = next;
        }

        nextAfterContinue = offset;
        return ms.ToArray();
    }

    private static List<string> ParseSst(byte[] payload)
    {
        var result = new List<string>();
        if (payload.Length < 8)
        {
            return result;
        }

        int uniqueCount = BinaryPrimitives.ReadInt32LittleEndian(payload.AsSpan(4, 4));
        int offset = 8;
        for (int i = 0; i < uniqueCount && offset < payload.Length; i++)
        {
            if (offset + 3 > payload.Length)
            {
                break;
            }

            var charCount = ReadUInt16(payload, offset);
            offset += 2;
            var flags = payload[offset++];
            bool highByte = (flags & 0x01) != 0;
            bool hasRich = (flags & 0x08) != 0;
            bool hasExt = (flags & 0x04) != 0;

            ushort richRunCount = 0;
            uint extSize = 0;
            if (hasRich)
            {
                if (offset + 2 > payload.Length) break;
                richRunCount = ReadUInt16(payload, offset);
                offset += 2;
            }
            if (hasExt)
            {
                if (offset + 4 > payload.Length) break;
                extSize = BinaryPrimitives.ReadUInt32LittleEndian(payload.AsSpan(offset, 4));
                offset += 4;
            }

            var byteCount = charCount * (highByte ? 2 : 1);
            if (offset + byteCount > payload.Length)
            {
                break;
            }

            var textBytes = payload.AsSpan(offset, byteCount);
            offset += byteCount;
            var text = highByte
                ? Encoding.Unicode.GetString(textBytes)
                : Encoding.GetEncoding(28591).GetString(textBytes);
            result.Add(text);

            var richBytes = richRunCount * 4;
            if (offset + richBytes > payload.Length) break;
            offset += richBytes;

            if (offset + (int)extSize > payload.Length) break;
            offset += (int)extSize;
        }
        return result;
    }

    private static BoundSheet ParseBoundSheet(ReadOnlySpan<byte> payload)
    {
        if (payload.Length < 8)
        {
            throw new FormatException("Invalid XLS file: BOUNDSHEET payload is too short.");
        }

        var bofOffset = (int)BinaryPrimitives.ReadUInt32LittleEndian(payload[..4]);
        var (name, _) = ParseShortUnicodeString(payload, 6);
        return new BoundSheet(name, bofOffset);
    }

    private static (string Value, int BytesRead) ParseShortUnicodeString(ReadOnlySpan<byte> source, int start)
    {
        if (start + 2 > source.Length)
        {
            return (string.Empty, 0);
        }

        var charCount = source[start];
        var flags = source[start + 1];
        bool highByte = (flags & 0x01) != 0;

        var offset = start + 2;
        var byteCount = charCount * (highByte ? 2 : 1);
        if (offset + byteCount > source.Length)
        {
            byteCount = Math.Max(0, source.Length - offset);
        }

        var text = highByte
            ? Encoding.Unicode.GetString(source.Slice(offset, byteCount))
            : Encoding.GetEncoding(28591).GetString(source.Slice(offset, byteCount));
        return (text, 2 + byteCount);
    }

    private static void SetCell(XlWorksheet ws, int row, int col, string value)
    {
        while (ws.Rows.Count <= row)
        {
            ws.Rows.Add(new XlRow(ws));
        }

        var xlRow = ws.Rows[row];
        while (xlRow.Cells.Count <= col)
        {
            xlRow.Cells.Add(new XlCell(xlRow) { Value = string.Empty });
        }
        xlRow.Cells[col].Value = value;
    }

    private static void SetFormulaCell(XlWorksheet ws, int row, int col, string formula, string value)
    {
        while (ws.Rows.Count <= row)
        {
            ws.Rows.Add(new XlRow(ws));
        }

        var xlRow = ws.Rows[row];
        while (xlRow.Cells.Count <= col)
        {
            xlRow.Cells.Add(new XlCell(xlRow) { Value = string.Empty });
        }
        xlRow.Cells[col].SetFormula(formula, value);
    }

    private static void ParseFormulaRecord(XlWorksheet ws, ReadOnlySpan<byte> payload,
        Dictionary<(int, int, int, int), byte[]> sharedFormulas)
    {
        if (payload.Length < 20) return;

        var row = ReadUInt16(payload, 0);
        var col = ReadUInt16(payload, 2);
        var grbit = ReadUInt16(payload, 14);
        var cachedValue = ParseCachedValue(payload.Slice(6, 8));

        int rpnStart = 20;
        if (rpnStart >= payload.Length) return;
        int rpnLen = payload.Length - rpnStart;

        // Check if this is a shared formula child (references a master)
        bool isShared = (grbit & 0x08) != 0;

        if (isShared && rpnLen == 0)
        {
            // This cell references a shared formula master — try to resolve
            // The chn field at bytes 16-19 contains the index, but matching to masters is complex.
            // For now, skip child shared formula cells without RPN data.
            return;
        }

        var formulaBytes = payload.Slice(rpnStart, rpnLen).ToArray();

        // Try to decompile via the pluggable hook
        var decompiler = XlHelper.Biff8FormulaDecompiler;
        if (decompiler is not null)
        {
            try
            {
                var formula = decompiler(formulaBytes);
                if (!string.IsNullOrEmpty(formula))
                {
                    SetFormulaCell(ws, row, col, formula, cachedValue);
                    return;
                }
            }
            catch
            {
                // Fall through to value-only
            }
        }

        // No decompiler — store cached value only
        if (!string.IsNullOrEmpty(cachedValue))
            SetCell(ws, row, col, cachedValue);
    }

    private static void ParseSharedFormulaRecord(ReadOnlySpan<byte> payload,
        Dictionary<(int, int, int, int), byte[]> sharedFormulas)
    {
        if (payload.Length < 8) return;

        var rwFirst = ReadUInt16(payload, 0);
        var rwLast = ReadUInt16(payload, 2);
        var colFirst = payload[4];
        var colLast = payload[5];
        // Skip 2 reserved bytes at offset 6-7
        var key = (rwFirst, rwLast, colFirst, colLast);

        int rpnStart = 8;
        if (rpnStart < payload.Length)
        {
            int rpnLen = payload.Length - rpnStart;
            sharedFormulas[key] = payload.Slice(rpnStart, rpnLen).ToArray();
        }
    }

    private static string ParseCachedValue(ReadOnlySpan<byte> result)
    {
        if (result.Length < 6) return "";

        // BIFF8 cached value encoding:
        // Empty/blank: first byte = 0x03, or all zeros
        // Number: IEEE 754 double (standard encoding)
        // String: not stored in 8 bytes in BIFF8 (separate STRING record follows)
        // Bool: first byte = 0x01 (or 0x04), value in second byte
        // Error: first byte = 0x02 (or 0x05), error code in second byte

        byte type = result[0];

        if (type == 0x01 || type == 0x04)
        {
            // Boolean
            return result[1] != 0 ? "TRUE" : "FALSE";
        }
        if (type == 0x02 || type == 0x05)
        {
            // Error
            byte code = result[1];
            return code switch
            {
                0x00 => "#NULL!",
                0x07 => "#DIV/0!",
                0x0F => "#VALUE!",
                0x17 => "#REF!",
                0x1D => "#NAME?",
                0x24 => "#NUM!",
                0x2A => "#N/A",
                _ => "#ERR!",
            };
        }
        if (type == 0x00)
        {
            // String (in BIFF8, empty string in cached value)
            return "";
        }

        // Treat as a number (IEEE 754 double)
        try
        {
            double num = BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(result));
            if (!double.IsNaN(num) && !double.IsInfinity(num))
                return num.ToString("G15");
        }
        catch { }

        return "";
    }

    private static double DecodeRk(uint rk)
    {
        bool divideBy100 = (rk & 0x01) != 0;
        bool isInteger = (rk & 0x02) != 0;

        double value;
        if (isInteger)
        {
            value = unchecked((int)rk) >> 2;
        }
        else
        {
            var raw = (long)(rk & 0xFFFFFFFC) << 32;
            value = BitConverter.Int64BitsToDouble(raw);
        }

        return divideBy100 ? value / 100d : value;
    }

    private static bool TryReadRecord(
        ReadOnlySpan<byte> stream,
        int offset,
        out ushort recordId,
        out int payloadOffset,
        out int payloadLength,
        out int nextOffset)
    {
        recordId = 0;
        payloadOffset = 0;
        payloadLength = 0;
        nextOffset = offset;

        if (offset < 0 || offset + 4 > stream.Length)
        {
            return false;
        }

        recordId = ReadUInt16(stream, offset);
        payloadLength = ReadUInt16(stream, offset + 2);
        payloadOffset = offset + 4;
        nextOffset = payloadOffset + payloadLength;
        return nextOffset <= stream.Length;
    }

    private static ushort ReadUInt16(ReadOnlySpan<byte> source, int offset) =>
        BinaryPrimitives.ReadUInt16LittleEndian(source.Slice(offset, 2));

    private sealed record BoundSheet(string Name, int BofOffset);
}
