using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Rules;
using SPOReplaceLink.OpenXml.Parts;
using static SPOReplaceLink.OpenXml.Tests.TestWorkbook;

namespace SPOReplaceLink.OpenXml.Tests;

/// <summary>OOX-6 (xl/connections.xml) and OOX-8 (Power Query in customXml → DataMashup).</summary>
public sealed class DataSourceTests : IDisposable
{
    private const string Connections = "xl/connections.xml";
    private const string Item = "customXml/item1.xml";
    private const string New = Host + "/sites/finance";

    private readonly string _dir = Path.Combine(Path.GetTempPath(), "spo-rl-tests", Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        if (Directory.Exists(_dir))
            Directory.Delete(_dir, true);
    }

    private static RuleSet Rules()
    {
        var build = RuleSet.Build(
        [
            new(Host + "/sites/old", New),
            new(@"\\fs01\Finance", @"\\fs02\Fin"),
        ], Host);
        Assert.True(build.IsValid, string.Join("\n", build.Issues));
        return build.RuleSet!;
    }

    private (FileResult Result, string Output) Run(TestWorkbook wb, PlacesOptions? places = null)
    {
        string input = wb.Save(Path.Combine(_dir, "in", Guid.NewGuid().ToString("N") + ".xlsx"));
        string output = Path.Combine(_dir, "out", Guid.NewGuid().ToString("N") + ".xlsx");
        var result = new WorkbookProcessor(Rules(), places ?? new PlacesOptions()).Process(input, output, null, "a.xlsx");
        Assert.Equal(result.OutputWritten, File.Exists(output));
        return (result, output);
    }

    private static TestWorkbook Base() => new TestWorkbook()
        .Set("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>")
        .Set("xl/workbook.xml", $"<workbook xmlns=\"{NsMain}\"/>");

    private const string ConnectionsXml =
        $"<connections xmlns=\"{NsMain}\">" +
        "<connection id=\"1\" name=\"fin\" type=\"1\" odcFile=\"\\\\fs01\\Finance\\Data\\fin.odc\" refreshedVersion=\"8\">" +
        "<dbPr connection=\"DSN=Excel Files;DBQ=\\\\fs01\\Finance\\Data\\Звіт 1.xlsx;DefaultDir=\\\\fs01\\Finance\\Data;DriverId=1046;\" " +
        "command=\"SELECT *&#13;&#10;FROM `\\\\fs01\\Finance\\Data\\Звіт 1.xlsx`.`Аркуш1$` WHERE x = 'a''b'\"/></connection>" +
        "<connection id=\"2\" name=\"ace\" type=\"5\"><dbPr connection=\"Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;" +
        "Data Source=&quot;\\\\fs01\\Finance\\a;b.accdb&quot;;Extended Properties=&quot;Excel 12.0 Xml;HDR=YES&quot;\" command=\"Table1\" commandType=\"3\"/></connection>" +
        $"<connection id=\"3\" name=\"web\" type=\"4\"><webPr url=\"{Host}/sites/old/Lists/x/AllItems.aspx?a=1&amp;b=2\" htmlTables=\"1\"/></connection>" +
        "<connection id=\"4\" name=\"Query - Q1\" type=\"5\" description=\"\\\\fs01\\Finance\\left as is\"><dbPr connection=\"Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Q1;\" command=\"SELECT * FROM [Q1]\"/></connection>" +
        "</connections>";

    [Fact]
    public void RewritesConnectionStringsCommandsAndFiles()
    {
        var (result, output) = Run(Base().Set(Connections, ConnectionsXml));

        Assert.Equal(FileStatus.Planned, result.Status);
        Assert.Equal(6, result.ReplacementsOf(LinkKind.Connection));
        string xml = ReadPart(output, Connections);
        const string lib = @"\\fs02\Fin";
        Assert.Contains($"odcFile=\"{lib}\\Data\\fin.odc\"", xml);
        Assert.Contains($"DBQ={lib}\\Data\\Звіт 1.xlsx;DefaultDir={lib}\\Data;DriverId=1046;", xml);
        // The line break keeps its character references; the SQL literal with '' is not touched.
        Assert.Contains($"command=\"SELECT *&#13;&#10;FROM `{lib}\\Data\\Звіт 1.xlsx`.`Аркуш1$` WHERE x = 'a''b'\"", xml);
        Assert.Contains($"Data Source=&quot;{lib}\\a;b.accdb&quot;;Extended Properties=&quot;Excel 12.0 Xml;HDR=YES&quot;", xml);
        Assert.Contains($"url=\"{New}/Lists/x/AllItems.aspx?a=1&amp;b=2\"", xml);
        // Power Query connection has no links; a description is free text and stays as a residual.
        Assert.Contains("Location=Q1;\" command=\"SELECT * FROM [Q1]\"", xml);
        Assert.Single(result.Residuals);
        Assert.Contains("left as is", result.Residuals[0].Context);
    }

    [Fact]
    public void DisabledConnectionsAreLeftAndReported()
    {
        var (result, _) = Run(Base().Set(Connections, ConnectionsXml), new PlacesOptions { Connections = false });

        Assert.Equal(FileStatus.Unchanged, result.Status);
        Assert.Contains(result.Links, l => l.Kind == LinkKind.Connection && l.Action == LinkAction.PlaceDisabled);
        Assert.True(result.Residuals.Count >= 6);
    }

    [Theory]
    [InlineData("Data Source=x;", "x")]
    [InlineData("A = 'q''w' ; B={a}}b};C=\"z\"", "q''w|a}}b|z")]
    [InlineData("Provider=P;Data Source= \\\\s\\a b.xlsx  ;", "P|\\\\s\\a b.xlsx")]
    public void ParsesConnectionStringValues(string s, string expected)
    {
        var values = PartRewriters.ConnectionStringValues(s).Select(v => s.Substring(v.Start, v.Length));
        Assert.Equal(expected, string.Join("|", values));
    }

    private const string Section =
        "section Section1;\r\n" +
        "\r\n" +
        "shared Q1 = let\r\n" +
        $"    // was {Host}/sites/old/Shared Documents\r\n" +
        $"    Source = Excel.Workbook(Web.Contents(\"{Host}/sites/old/Shared Documents/Звіт.xlsx\"), null, true),\r\n" +
        $"    Files = SharePoint.Files(\"{Host}/sites/old\", [ApiVersion = 15]),\r\n" +
        "    Csv = Csv.Document(File.Contents(\"\\\\fs01\\Finance\\a \"\"b\"\".csv\")),\r\n" +
        $"    #\"{Host}/sites/old step\" = Text.Split(\"a/b\", \"/\"),\r\n" +
        $"    Note = \"see {Host}/sites/old/x.xlsx please\"\r\n" +
        "in\r\n" +
        "    Source;";

    private static readonly byte[] Permissions = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?><PermissionList/>");
    private static readonly byte[] Bindings = [1, 2, 3, 4, 5];

    /// <summary>A DataMashup the way Excel writes it: [MS-QDEFF] blocks, base64 in a UTF-16 customXml part.</summary>
    internal static byte[] MashupPart(string m)
    {
        using var zipStream = new MemoryStream();
        using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, text) in new[]
            {
                ("[Content_Types].xml", "<Types/>"),
                ("Config/Package.xml", "<Package/>"),
                ("Formulas/Section1.m", m),
            })
            {
                using var s = zip.CreateEntry(name).Open();
                s.Write(Encoding.UTF8.GetPreamble());
                s.Write(Encoding.UTF8.GetBytes(text));
            }
        }
        byte[] metadataXml = Encoding.UTF8.GetBytes("<LocalPackageMetadataFile/>");
        byte[] metadata = [.. U32(0), .. U32(metadataXml.Length), .. metadataXml, .. U32(0)];
        byte[] package = zipStream.ToArray();
        byte[] data = [.. U32(0), .. U32(package.Length), .. package, .. U32(Permissions.Length), .. Permissions,
            .. U32(metadata.Length), .. metadata, .. U32(Bindings.Length), .. Bindings];
        string xml = "<?xml version=\"1.0\" encoding=\"utf-16\"?><DataMashup xmlns=\"http://schemas.microsoft.com/DataMashup\">"
            + Convert.ToBase64String(data) + "</DataMashup>";
        return [.. Encoding.Unicode.GetPreamble(), .. Encoding.Unicode.GetBytes(xml)];
    }

    private static byte[] U32(int v)
    {
        var b = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(b, v);
        return b;
    }

    /// <summary>Formulas/Section1.m and the bytes after the package block of a written mashup part.</summary>
    private static (string M, byte[] Tail) ReadMashup(byte[] part)
    {
        Assert.True(part[0] == 0xFF && part[1] == 0xFE, "UTF-16 with BOM is kept");
        string xml = Encoding.Unicode.GetString(part, 2, part.Length - 2);
        int start = xml.IndexOf('>', xml.IndexOf("<DataMashup", StringComparison.Ordinal)) + 1;
        byte[] data = Convert.FromBase64String(xml[start..xml.IndexOf("</DataMashup>", StringComparison.Ordinal)]);
        int length = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(4));
        using var zip = new ZipArchive(new MemoryStream(data, 8, length), ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("Formulas/Section1.m")!.Open(), Encoding.UTF8);
        return (reader.ReadToEnd(), data[(8 + length)..]);
    }

    [Fact]
    public void RewritesPowerQueryFormulas()
    {
        var (result, output) = Run(Base().Set(Item, MashupPart(Section)));

        Assert.Equal(FileStatus.Planned, result.Status);
        Assert.Equal(4, result.ReplacementsOf(LinkKind.PowerQuery));
        Assert.Equal([Item], result.ChangedParts);
        var (m, rest) = ReadMashup(ReadParts(output)[Item]);
        Assert.Contains($"Web.Contents(\"{New}/Shared Documents/Звіт.xlsx\")", m);
        Assert.Contains($"SharePoint.Files(\"{New}\", [ApiVersion = 15])", m);
        Assert.Contains("File.Contents(\"\\\\fs02\\Fin\\a \"\"b\"\".csv\")", m);
        Assert.Contains($"Note = \"see {New}/x.xlsx please\"", m);
        // Comments and quoted identifiers are not text; "/" is not a server-relative link.
        Assert.Contains($"// was {Host}/sites/old/Shared Documents", m);
        Assert.Contains($"#\"{Host}/sites/old step\"", m);
        Assert.DoesNotContain(result.Links, l => l.OldValue == "/");
        // Permissions, metadata and bindings are copied as they are.
        var (_, before) = ReadMashup(MashupPart(Section));
        Assert.Equal(before, rest);
        Assert.Equal(2, result.Residuals.Count);
        Assert.All(result.Residuals, r => Assert.Equal(Item + "!Formulas/Section1.m", r.Part));
    }

    [Fact]
    public void SecondPassChangesNothing()
    {
        var (_, output) = Run(Base().Set(Item, MashupPart(Section)).Set(Connections, ConnectionsXml));
        var wb = new TestWorkbook();
        foreach (var (name, bytes) in ReadParts(output))
            wb.Set(name, bytes);
        var (again, _) = Run(wb);
        Assert.Equal(FileStatus.Unchanged, again.Status);
        Assert.Equal(0, again.Replacements);
        Assert.Contains(again.Links, l => l.Kind == LinkKind.PowerQuery && l.OldValue.StartsWith(New, StringComparison.Ordinal));
    }

    [Fact]
    public void DisabledPowerQueryIsLeftAndReported()
    {
        var (result, _) = Run(Base().Set(Item, MashupPart(Section)), new PlacesOptions { PowerQuery = false });

        Assert.Equal(FileStatus.Unchanged, result.Status);
        Assert.Contains(result.Links, l => l.Kind == LinkKind.PowerQuery && l.Action == LinkAction.PlaceDisabled);
        Assert.Contains(result.Residuals, r => r.Context.Contains("Web.Contents"));
    }

    [Fact]
    public void OtherCustomXmlIsOnlyScanned()
    {
        var (result, _) = Run(Base().Set(Item, $"<root><a>{Host}/sites/old/x</a></root>"));
        Assert.Equal(FileStatus.Unchanged, result.Status);
        Assert.Single(result.Residuals);
    }

    [Fact]
    public void BrokenMashupFailsTheFile()
    {
        var (result, _) = Run(Base().Set(Item, "<DataMashup xmlns=\"http://schemas.microsoft.com/DataMashup\">AAAA</DataMashup>"));
        Assert.Equal("failed:file", result.StatusCode);
    }
}
