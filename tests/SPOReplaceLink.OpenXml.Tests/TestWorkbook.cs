using System.IO.Compression;
using System.Text;

namespace SPOReplaceLink.OpenXml.Tests;

/// <summary>
/// Minimal but valid SpreadsheetML package with links in every supported place.
/// Tests change single parts to cover edge cases; Excel opens the default package without repair.
/// </summary>
internal sealed class TestWorkbook
{
    public const string Host = "https://contoso.sharepoint.com";
    public const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string Decl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

    public Dictionary<string, byte[]> Parts { get; } = new(StringComparer.Ordinal);

    public static TestWorkbook Default()
    {
        var wb = new TestWorkbook();
        wb.Set("[Content_Types].xml", Decl +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            "<Override PartName=\"/xl/externalLinks/externalLink1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml\"/>" +
            "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
            "</Types>");
        wb.Set("_rels/.rels", Decl +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>");
        wb.Set("xl/workbook.xml", Decl +
            $"<workbook xmlns=\"{NsMain}\" xmlns:r=\"{NsRel}\">" +
            "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
            "<externalReferences><externalReference r:id=\"rId2\"/></externalReferences>" +
            "</workbook>");
        wb.Set("xl/_rels/workbook.xml.rels", Decl +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink\" Target=\"externalLinks/externalLink1.xml\"/>" +
            "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
            "</Relationships>");
        wb.Set("xl/worksheets/sheet1.xml", Decl +
            $"<worksheet xmlns=\"{NsMain}\" xmlns:r=\"{NsRel}\">\r\n" +
            "  <sheetData>\r\n" +
            "    <row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"inlineStr\"><is><t>rel</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>internal</t></is></c></row>\r\n" +
            $"    <row r=\"2\"><c r=\"A2\" t=\"str\"><f>HYPERLINK(\"{Host}/Document/Fin/Звіт 1.xlsx\",\"Звіт\")</f><v>Звіт</v></c></row>\r\n" +
            "    <row r=\"3\"><c r=\"A3\"><f>[1]Sheet1!A1</f><v>42</v></c></row>\r\n" +
            $"    <row r=\"4\"><c r=\"A4\" t=\"str\"><f>\"{Host}/Document/Analytics/\"&amp;\"x.xlsx\"</f><v>{Host}/Document/Analytics/x.xlsx</v></c></row>\r\n" +
            "  </sheetData>\r\n" +
            "  <hyperlinks><hyperlink ref=\"A1\" r:id=\"rId1\"/><hyperlink ref=\"B1\" r:id=\"rId2\"/><hyperlink ref=\"C1\" location=\"Sheet1!A3\" display=\"internal\"/></hyperlinks>\r\n" +
            "</worksheet>");
        wb.Set("xl/worksheets/_rels/sheet1.xml.rels", Decl +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id=\"rId1\" Type=\"{NsRel}/hyperlink\" Target=\"{Host}/Document/Fin/a%20b.xlsx?web=1&amp;x=2\" TargetMode=\"External\"/>" +
            $"<Relationship Id=\"rId2\" Type=\"{NsRel}/hyperlink\" Target=\"../Analytics/rel.xlsx\" TargetMode=\"External\"/>" +
            "</Relationships>");
        wb.Set("xl/externalLinks/externalLink1.xml", Decl +
            $"<externalLink xmlns=\"{NsMain}\" xmlns:r=\"{NsRel}\">" +
            "<externalBook r:id=\"rId1\"><sheetNames><sheetName val=\"Sheet1\"/></sheetNames>" +
            "<sheetDataSet><sheetData sheetId=\"0\"><row r=\"1\"><cell r=\"A1\"><v>42</v></cell></row></sheetData></sheetDataSet>" +
            "</externalBook></externalLink>");
        wb.Set("xl/externalLinks/_rels/externalLink1.xml.rels", Decl +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id=\"rId1\" Type=\"{NsRel}/externalLinkPath\" Target=\"/Document/Analytics/Source.xlsx\" TargetMode=\"External\"/>" +
            "</Relationships>");
        wb.Set("xl/sharedStrings.xml", Decl +
            $"<sst xmlns=\"{NsMain}\" count=\"1\" uniqueCount=\"1\"><si><t>link</t></si></sst>");
        return wb;
    }

    public TestWorkbook Set(string name, string text)
    {
        Parts[name] = Encoding.UTF8.GetBytes(text);
        return this;
    }

    public TestWorkbook Set(string name, byte[] bytes)
    {
        Parts[name] = bytes;
        return this;
    }

    public string Text(string name) => Encoding.UTF8.GetString(Parts[name]);

    public string Save(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var fs = File.Create(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Create);
        foreach (var (name, bytes) in Parts)
        {
            var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using var s = entry.Open();
            s.Write(bytes);
        }
        return path;
    }

    /// <summary>
    /// Saves every part deflated the way Excel does, an empty one as a two-byte deflate stream
    /// (ZipArchive writes an empty entry as stored).
    /// </summary>
    public string SaveLikeExcel(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var fs = File.Create(path);
        using var w = new BinaryWriter(fs);
        var central = new List<(byte[] Name, uint Crc, int Compressed, int Length, long Offset)>();
        foreach (var (name, bytes) in Parts)
        {
            byte[] data = bytes.Length == 0 ? [0x03, 0x00] : Deflate(bytes);
            byte[] n = Encoding.UTF8.GetBytes(name);
            uint crc = Crc32(bytes);
            central.Add((n, crc, data.Length, bytes.Length, fs.Position));
            w.Write(0x04034b50u);
            Header(w, crc, data.Length, bytes.Length, n.Length);
            w.Write((ushort)0); // extra
            w.Write(n);
            w.Write(data);
        }
        long start = fs.Position;
        foreach (var e in central)
        {
            w.Write(0x02014b50u);
            w.Write((ushort)20); // made by
            Header(w, e.Crc, e.Compressed, e.Length, e.Name.Length);
            w.Write(new byte[12]); // extra, comment, disk, internal and external attributes
            w.Write((uint)e.Offset);
            w.Write(e.Name);
        }
        long end = fs.Position;
        w.Write(0x06054b50u);
        w.Write(0u); // disks
        w.Write((ushort)central.Count);
        w.Write((ushort)central.Count);
        w.Write((uint)(end - start));
        w.Write((uint)start);
        w.Write((ushort)0);
        return path;
    }

    private static void Header(BinaryWriter w, uint crc, int compressed, int length, int nameLength)
    {
        w.Write((ushort)20); // version needed
        w.Write((ushort)0x800); // UTF-8 names
        w.Write((ushort)8); // deflate
        w.Write((ushort)0); // time
        w.Write((ushort)0x21); // 1980-01-01
        w.Write(crc);
        w.Write(compressed);
        w.Write(length);
        w.Write((ushort)nameLength);
    }

    private static byte[] Deflate(byte[] bytes)
    {
        using var ms = new MemoryStream();
        using (var d = new DeflateStream(ms, CompressionLevel.Optimal, leaveOpen: true))
            d.Write(bytes);
        return ms.ToArray();
    }

    private static uint Crc32(byte[] bytes)
    {
        uint crc = 0xFFFFFFFF;
        foreach (byte b in bytes)
        {
            crc ^= b;
            for (int k = 0; k < 8; k++)
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
        }
        return ~crc;
    }

    public static Dictionary<string, byte[]> ReadParts(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var e in zip.Entries)
        {
            using var s = e.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            parts[e.FullName] = ms.ToArray();
        }
        return parts;
    }

    public static string ReadPart(string path, string name) => Encoding.UTF8.GetString(ReadParts(path)[name]);
}
