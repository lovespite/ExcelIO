using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace ExcelIO;

internal static class XlsCompoundReader
{
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSect = 0xFFFFFFFF;
    private const uint FatSect = 0xFFFFFFFD;

    public static ReadOnlySpan<byte> ReadWorkbookBytes(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);

        return ms.ToArray();
    }

    public static ReadOnlySpan<byte> ReadWorkbookStream(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 512)
        {
            throw new FormatException("Invalid XLS file: CFB header is too short.");
        }

        ReadOnlySpan<byte> header = bytes[..512];
        var signatureLo = BinaryPrimitives.ReadUInt64LittleEndian(header[..8]);
        if (signatureLo != 0xE11AB1A1E011CFD0UL)
        {
            throw new FormatException("Invalid XLS file: CFB signature mismatch.");
        }

        var majorVersion = ReadUInt16(header, 0x1A);
        var sectorShift = ReadUInt16(header, 0x1E);
        var miniSectorShift = ReadUInt16(header, 0x20);
        var sectorSize = 1 << sectorShift;
        var miniSectorSize = 1 << miniSectorShift;
        if (sectorSize <= 0 || miniSectorSize <= 0)
        {
            throw new FormatException("Invalid XLS file: CFB sector size is invalid.");
        }

        var numFatSectors = ReadUInt32(header, 0x2C);
        var firstDirSector = ReadUInt32(header, 0x30);
        var miniCutoff = ReadUInt32(header, 0x38);
        var firstMiniFatSector = ReadUInt32(header, 0x3C);
        var numMiniFatSectors = ReadUInt32(header, 0x40);
        var firstDifatSector = ReadUInt32(header, 0x44);
        var numDifatSectors = ReadUInt32(header, 0x48);

        var difat = new List<uint>((int)Math.Max(0, numFatSectors));
        for (int i = 0; i < 109; i++)
        {
            var sid = ReadUInt32(header, 0x4C + i * 4);
            if (sid != FreeSect)
            {
                difat.Add(sid);
            }
        }

        if (numDifatSectors > 0 && firstDifatSector != EndOfChain)
        {
            var entriesPerDifatSector = (sectorSize / 4) - 1;
            var currentDifat = firstDifatSector;
            for (uint i = 0; i < numDifatSectors && currentDifat != EndOfChain; i++)
            {
                var difatBytes = ReadSector(bytes, currentDifat, sectorSize);
                for (int j = 0; j < entriesPerDifatSector; j++)
                {
                    var sid = BinaryPrimitives.ReadUInt32LittleEndian(difatBytes.Slice(j * 4, 4));
                    if (sid != FreeSect)
                    {
                        difat.Add(sid);
                    }
                }
                currentDifat = BinaryPrimitives.ReadUInt32LittleEndian(difatBytes.Slice(entriesPerDifatSector * 4, 4));
            }
        }

        if (difat.Count == 0)
        {
            throw new FormatException("Invalid XLS file: no FAT sectors found.");
        }
        if (numFatSectors > 0 && difat.Count > numFatSectors)
        {
            difat = difat.Take((int)numFatSectors).ToList();
        }

        var fat = ReadFat(bytes, difat, sectorSize);
        var directoryStream = ReadStreamByFat(bytes, fat, firstDirSector, null, sectorSize);
        var entries = ParseDirectoryEntries(directoryStream, majorVersion);
        var root = entries.FirstOrDefault(x => x.Type == 5)
            ?? throw new FormatException("Invalid XLS file: Root Entry not found.");
        var workbook = entries.FirstOrDefault(x =>
            x.Type == 2 && (string.Equals(x.Name, "Workbook", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(x.Name, "Book", StringComparison.OrdinalIgnoreCase)))
            ?? throw new FormatException("Invalid XLS file: Workbook stream not found.");

        if (workbook.Size < miniCutoff)
        {
            if (firstMiniFatSector == EndOfChain || numMiniFatSectors == 0)
            {
                throw new FormatException("Invalid XLS file: MiniFAT is required but missing.");
            }

            var miniFat = ReadMiniFat(bytes, fat, firstMiniFatSector, numMiniFatSectors, sectorSize);
            var miniStream = ReadStreamByFat(bytes, fat, root.StartSector, (int)root.Size, sectorSize);
            return ReadStreamByMiniFat(miniStream, miniFat, workbook.StartSector, (int)workbook.Size, miniSectorSize);
        }

        return ReadStreamByFat(bytes, fat, workbook.StartSector, (int)workbook.Size, sectorSize);
    }

    private static int[] ReadFat(ReadOnlySpan<byte> bytes, IReadOnlyList<uint> fatSectors, int sectorSize)
    {
        var entries = new List<int>(fatSectors.Count * (sectorSize / 4));
        foreach (var fatSector in fatSectors)
        {
            var sector = ReadSector(bytes, fatSector, sectorSize);
            for (int i = 0; i < sectorSize; i += 4)
            {
                entries.Add(BinaryPrimitives.ReadInt32LittleEndian(sector.Slice(i, 4)));
            }
        }
        return [.. entries];
    }

    private static int[] ReadMiniFat(
        ReadOnlySpan<byte> bytes,
        int[] fat,
        uint firstMiniFatSector,
        uint numMiniFatSectors,
        int sectorSize)
    {
        var miniFatBytes = ReadStreamByFat(bytes, fat, firstMiniFatSector, (int)(numMiniFatSectors * (uint)sectorSize), sectorSize);
        var miniFat = new int[miniFatBytes.Length / 4];
        for (int i = 0; i < miniFat.Length; i++)
        {
            miniFat[i] = BinaryPrimitives.ReadInt32LittleEndian(miniFatBytes.Slice(i * 4, 4));
        }
        return miniFat;
    }

    private static ReadOnlySpan<byte> ReadStreamByFat(ReadOnlySpan<byte> file, int[] fat, uint startSector, int? expectedSize, int sectorSize)
    {
        if (startSector == EndOfChain)
        {
            return [];
        }

        using var ms = new MemoryStream();
        var visited = new HashSet<uint>();
        var current = startSector;
        while (current != EndOfChain)
        {
            if (!visited.Add(current))
            {
                throw new FormatException("Invalid XLS file: FAT chain loop detected.");
            }

            var sector = ReadSector(file, current, sectorSize);
            ms.Write(sector);

            if (expectedSize.HasValue && ms.Length >= expectedSize.Value)
            {
                break;
            }

            if (current >= fat.Length)
            {
                throw new FormatException("Invalid XLS file: FAT chain index is out of range.");
            }

            current = unchecked((uint)fat[current]);
        }

        var data = ms.ToArray();
        if (expectedSize.HasValue && expectedSize.Value < data.Length)
        {
            Array.Resize(ref data, expectedSize.Value);
        }
        return data;
    }

    private static ReadOnlySpan<byte> ReadStreamByMiniFat(
        ReadOnlySpan<byte> miniStream,
        int[] miniFat,
        uint startMiniSector,
        int expectedSize,
        int miniSectorSize)
    {
        if (expectedSize == 0 || startMiniSector == EndOfChain)
        {
            return [];
        }

        using var ms = new MemoryStream();
        var visited = new HashSet<uint>();
        var current = startMiniSector;
        while (current != EndOfChain && ms.Length < expectedSize)
        {
            if (!visited.Add(current))
            {
                throw new FormatException("Invalid XLS file: MiniFAT chain loop detected.");
            }

            var offset = checked((int)current * miniSectorSize);
            if (offset < 0 || offset + miniSectorSize > miniStream.Length)
            {
                throw new FormatException("Invalid XLS file: MiniFAT sector out of range.");
            }

            ms.Write(miniStream.Slice(offset, miniSectorSize));
            if (current >= miniFat.Length)
            {
                throw new FormatException("Invalid XLS file: MiniFAT index out of range.");
            }
            current = unchecked((uint)miniFat[current]);
        }

        var data = ms.ToArray();
        if (expectedSize < data.Length)
        {
            Array.Resize(ref data, expectedSize);
        }
        return data;
    }

    private static ReadOnlySpan<byte> ReadSector(ReadOnlySpan<byte> bytes, uint sectorId, int sectorSize)
    {
        var offset = checked((int)(sectorId + 1) * sectorSize);
        if (offset < 0 || offset + sectorSize > bytes.Length)
        {
            throw new FormatException("Invalid XLS file: sector offset out of range.");
        }

        //var buf = new byte[sectorSize];
        //bytes.Slice(offset, sectorSize).CopyTo(buf);
        //return buf;

        return bytes.Slice(offset, sectorSize);
    }

    private static List<DirectoryEntry> ParseDirectoryEntries(ReadOnlySpan<byte> directoryStream, ushort majorVersion)
    {
        var list = new List<DirectoryEntry>();
        for (int offset = 0; offset + 128 <= directoryStream.Length; offset += 128)
        {
            ReadOnlySpan<byte> entry = directoryStream.Slice(offset, 128);
            var nameLength = ReadUInt16(entry, 64);
            if (nameLength < 2 || nameLength > 64)
            {
                continue;
            }

            var charBytes = nameLength - 2;
            var name = Encoding.Unicode.GetString(entry.Slice(0, charBytes));
            var type = entry[66];
            var startSector = ReadUInt32(entry, 116);
            var size = majorVersion == 3
                ? ReadUInt32(entry, 120)
                : (long)ReadUInt64(entry, 120);

            list.Add(new DirectoryEntry(name.TrimEnd('\0'), type, startSector, size));
        }
        return list;
    }

    private static ushort ReadUInt16(ReadOnlySpan<byte> source, int offset) =>
        BinaryPrimitives.ReadUInt16LittleEndian(source.Slice(offset, 2));

    private static uint ReadUInt32(ReadOnlySpan<byte> source, int offset) =>
        BinaryPrimitives.ReadUInt32LittleEndian(source.Slice(offset, 4));

    private static ulong ReadUInt64(ReadOnlySpan<byte> source, int offset) =>
        BinaryPrimitives.ReadUInt64LittleEndian(source.Slice(offset, 8));

    private sealed record DirectoryEntry(string Name, byte Type, uint StartSector, long Size);
}
