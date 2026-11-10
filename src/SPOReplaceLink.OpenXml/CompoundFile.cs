using System.Buffers.Binary;
using System.Text;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.OpenXml;

/// <summary>
/// Just enough of an OLE compound file (MS-CFB) to list its streams and read the start of one, without Excel.
/// Anything malformed makes <see cref="Open"/> return null: the caller then leaves the file to Excel as before.
/// </summary>
internal sealed class CompoundFile : IDisposable
{
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint MaxSector = 0xFFFFFFFA;
    private const int EntrySize = 128;

    private readonly Stream _stream;
    private readonly int _sectorSize;
    private readonly uint _miniCutoff;
    private readonly uint _firstMiniFat;
    private readonly uint[] _fat;
    private readonly List<(string Name, int Type, uint Start, long Size)> _entries = [];

    private CompoundFile(Stream stream, byte[] header)
    {
        _stream = stream;
        int shift = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(0x1E));
        if (shift is not (9 or 12) || BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(0x20)) != 6)
            throw new InvalidDataException(L.T("Невідомий розмір сектора."));
        _sectorSize = 1 << shift;
        _miniCutoff = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(0x38));
        _firstMiniFat = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(0x3C));
        _fat = ReadFat(header);

        foreach (uint sector in Chain(BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(0x30))))
        {
            byte[] dir = ReadSector(sector);
            for (int i = 0; i + EntrySize <= dir.Length; i += EntrySize)
            {
                var e = dir.AsSpan(i, EntrySize);
                int length = BinaryPrimitives.ReadUInt16LittleEndian(e[0x40..]);
                string name = length >= 2 && length <= 64 ? Encoding.Unicode.GetString(e[..(length - 2)]) : string.Empty;
                long size = shift == 9 ? BinaryPrimitives.ReadUInt32LittleEndian(e[0x78..]) : BinaryPrimitives.ReadInt64LittleEndian(e[0x78..]);
                _entries.Add((name, e[0x42], BinaryPrimitives.ReadUInt32LittleEndian(e[0x74..]), size));
            }
        }
        if (_entries.Count == 0 || _entries[0].Type != 5)
            throw new InvalidDataException(L.T("Немає кореневого запису."));
    }

    public static CompoundFile? Open(string path)
    {
        var stream = File.OpenRead(path);
        try
        {
            byte[] header = new byte[512];
            if (stream.ReadAtLeast(header, header.Length, throwOnEndOfStream: false) < header.Length
                || !header.AsSpan(0, 8).SequenceEqual(PackageSniffer.CompoundSignature))
            {
                stream.Dispose();
                return null;
            }
            return new CompoundFile(stream, header);
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or ArgumentException or OverflowException or IndexOutOfRangeException)
        {
            stream.Dispose();
            return null;
        }
    }

    public bool Has(string name) => _entries.Any(e => e.Type == 2 && e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    /// <summary>Up to <paramref name="limit"/> bytes from the start of a stream; null if there is none or it cannot be read.</summary>
    public byte[]? Read(string name, int limit)
    {
        var entry = _entries.FirstOrDefault(e => e.Type == 2 && e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (entry.Name is null)
            return null;
        try
        {
            int length = (int)Math.Min(entry.Size, limit);
            var result = new byte[length];
            int done = 0;
            if (entry.Size < _miniCutoff)
            {
                // Small streams live in 64-byte sectors inside the mini stream, which is the root entry's chain.
                var miniFat = ReadTable(_firstMiniFat);
                var root = Chain(_entries[0].Start).ToList();
                foreach (uint mini in Chain(entry.Start, miniFat))
                {
                    if (done >= length)
                        break;
                    long offset = mini * 64L;
                    byte[] sector = ReadSector(root[(int)(offset / _sectorSize)]);
                    int n = Math.Min(64, length - done);
                    sector.AsSpan((int)(offset % _sectorSize), n).CopyTo(result.AsSpan(done));
                    done += n;
                }
            }
            else
            {
                foreach (uint sector in Chain(entry.Start))
                {
                    if (done >= length)
                        break;
                    int n = Math.Min(_sectorSize, length - done);
                    ReadSector(sector).AsSpan(0, n).CopyTo(result.AsSpan(done));
                    done += n;
                }
            }
            return done == length ? result : null;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or ArgumentException or IndexOutOfRangeException)
        {
            return null;
        }
    }

    private uint[] ReadFat(byte[] header)
    {
        uint count = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(0x2C));
        long maxSectors = _stream.Length / _sectorSize + 1;
        if (count > maxSectors)
            throw new InvalidDataException(L.T("Забагато секторів FAT."));
        var sectors = new List<uint>();
        for (int i = 0; i < 109 && sectors.Count < count; i++)
            sectors.Add(BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(0x4C + i * 4)));
        uint difat = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(0x44));
        for (int guard = 0; sectors.Count < count && difat <= MaxSector && guard < maxSectors; guard++)
        {
            byte[] d = ReadSector(difat);
            for (int i = 0; i < _sectorSize / 4 - 1 && sectors.Count < count; i++)
                sectors.Add(BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(i * 4)));
            difat = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(_sectorSize - 4));
        }
        if (sectors.Count < count)
            throw new InvalidDataException(L.T("Неповний DIFAT."));
        var fat = new uint[sectors.Count * (_sectorSize / 4)];
        for (int s = 0; s < sectors.Count; s++)
        {
            byte[] data = ReadSector(sectors[s]);
            for (int i = 0; i < _sectorSize / 4; i++)
                fat[s * (_sectorSize / 4) + i] = BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(i * 4));
        }
        return fat;
    }

    private uint[] ReadTable(uint start)
    {
        var table = new List<uint>();
        foreach (uint sector in Chain(start))
        {
            byte[] data = ReadSector(sector);
            for (int i = 0; i < _sectorSize / 4; i++)
                table.Add(BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(i * 4)));
        }
        return [.. table];
    }

    private IEnumerable<uint> Chain(uint start) => Chain(start, _fat);

    private static IEnumerable<uint> Chain(uint start, uint[] table)
    {
        // A loop in a damaged file never ends otherwise.
        for (int guard = 0; start != EndOfChain; guard++)
        {
            if (start > MaxSector || start >= table.Length || guard > table.Length)
                throw new InvalidDataException(L.T("Пошкоджений ланцюжок секторів."));
            yield return start;
            start = table[start];
        }
    }

    private byte[] ReadSector(uint sector)
    {
        long offset = (sector + 1L) * _sectorSize;
        if (offset >= _stream.Length)
            throw new InvalidDataException(L.T("Сектор за межами файлу."));
        var buffer = new byte[_sectorSize];
        _stream.Position = offset;
        // The last sector may be cut short; the rest stays zero.
        _stream.ReadAtLeast(buffer, buffer.Length, throwOnEndOfStream: false);
        return buffer;
    }

    public void Dispose() => _stream.Dispose();
}
