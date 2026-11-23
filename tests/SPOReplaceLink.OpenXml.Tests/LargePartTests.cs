using System.Text;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Rules;
using static SPOReplaceLink.OpenXml.Tests.TestWorkbook;

namespace SPOReplaceLink.OpenXml.Tests;

/// <summary>
/// Parts too big for one string (OutOfMemoryException on sheets of millions of rows) are read piece by piece;
/// the result must be the same as when the part is read whole.
/// </summary>
public sealed class LargePartTests : IDisposable
{
    private const string Sheet = "xl/worksheets/sheet1.xml";
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
            new("/Document/", Host + "/sites/finance/"),
            new("/Document/Analytics/", "/sites/finance/analytics_migrated/"),
        ], Host);
        Assert.True(build.IsValid, string.Join("\n", build.Issues));
        return build.RuleSet!;
    }

    private (FileResult Result, string Output) Run(string input, WorkbookProcessor processor)
    {
        string output = Path.Combine(_dir, "out", Guid.NewGuid().ToString("N") + Path.GetExtension(input));
        var result = processor.Process(input, output, Host + "/Document/Fin", Path.GetFileName(input));
        Assert.Equal(result.OutputWritten, File.Exists(output));
        string outDir = Path.GetDirectoryName(output)!;
        Assert.True(!Directory.Exists(outDir) || Directory.GetFiles(outDir, "*.part").Length == 0);
        return (result, output);
    }

    /// <summary>Every sheet, xml and bin part goes the piece-by-piece way, in pieces of a few dozen chars.</summary>
    private static WorkbookProcessor Streaming() => new(Rules(), new PlacesOptions(), largePart: -1, segmentChars: 48);

    private void AssertSameAsWhole(TestWorkbook wb, string ext = ".xlsx")
    {
        string input = wb.Save(Path.Combine(_dir, "in", Guid.NewGuid().ToString("N") + ext));
        var (whole, wholeOut) = Run(input, new WorkbookProcessor(Rules(), new PlacesOptions()));
        var (pieces, piecesOut) = Run(input, Streaming());

        Assert.Equal(whole.Status, pieces.Status);
        Assert.Equal(whole.Error, pieces.Error);
        Assert.Equal(whole.ChangedParts.Order(StringComparer.Ordinal), pieces.ChangedParts.Order(StringComparer.Ordinal));
        Assert.Equal(whole.Links.Select(l => (l.Part, l.Kind, l.OldValue, l.NewValue, l.Action)),
            pieces.Links.Select(l => (l.Part, l.Kind, l.OldValue, l.NewValue, l.Action)));
        Assert.Equal(whole.Residuals.Select(r => (r.Part, r.Match)).Order(), pieces.Residuals.Select(r => (r.Part, r.Match)).Order());
        if (whole.OutputWritten)
        {
            var a = ReadParts(wholeOut);
            var b = ReadParts(piecesOut);
            Assert.Equal(a.Keys, b.Keys);
            foreach (var (name, bytes) in a)
                Assert.True(bytes.AsSpan().SequenceEqual(b[name]), name);
        }
    }

    [Fact]
    public void LargeSheetIsRewrittenAsWhenReadWhole() => AssertSameAsWhole(Default());

    [Fact]
    public void LargeSheetWithPrefixesAndBomIsRewrittenAsWhenReadWhole()
    {
        var wb = Default();
        string sheet = wb.Text(Sheet)
            .Replace($"<worksheet xmlns=\"{NsMain}\"", $"<x:worksheet xmlns:x=\"{NsMain}\"")
            .Replace("</worksheet>", "</x:worksheet>")
            .Replace("<c ", "<x:c ").Replace("</c>", "</x:c>")
            .Replace("<f>", "<x:f>").Replace("</f>", "</x:f>")
            .Replace("<v>", "<x:v>").Replace("</v>", "</x:v>")
            .Replace("<row ", "<x:row ").Replace("</row>", "</x:row>")
            .Replace("<sheetData>", "<x:sheetData>").Replace("</sheetData>", "</x:sheetData>")
            .Replace("<hyperlinks>", "<x:hyperlinks>").Replace("</hyperlinks>", "</x:hyperlinks>")
            .Replace("<hyperlink ", "<x:hyperlink ")
            .Replace("<is><t>", "<x:is><x:t>").Replace("</t></is>", "</x:t></x:is>");
        wb.Set(Sheet, [.. Encoding.UTF8.GetPreamble(), .. Encoding.UTF8.GetBytes(sheet)]);
        AssertSameAsWhole(wb);
    }

    [Fact]
    public void ManyRowsWithLinksAreRewrittenAsWhenReadWhole()
    {
        var wb = Default();
        var rows = new StringBuilder();
        for (int i = 10; i < 400; i++)
        {
            rows.Append($"<row r=\"{i}\"><c r=\"A{i}\" t=\"str\"><f>HYPERLINK(\"{Host}/Document/Fin/f{i}.xlsx\")</f><v>x</v></c>")
                .Append($"<c r=\"B{i}\"><v>{i}</v></c></row>\r\n");
        }
        wb.Set(Sheet, wb.Text(Sheet).Replace("  </sheetData>", rows + "  </sheetData>"));
        AssertSameAsWhole(wb);
    }

    [Fact]
    public void UnchangedLargePartsStayUnchanged()
    {
        var wb = Default();
        // No links that the rules change anywhere.
        foreach (string name in wb.Parts.Keys.ToList())
            wb.Set(name, wb.Text(name).Replace("/Document/", "/Elsewhere/").Replace("../Analytics/", "https://other.example/"));
        wb.Set("xl/sharedStrings.xml", wb.Text("xl/sharedStrings.xml").Replace("<t>link</t>", $"<t>see {Host}/Document/Fin/old.xlsx</t>"));
        AssertSameAsWhole(wb);

        string input = wb.Save(Path.Combine(_dir, "in", "u.xlsx"));
        var (result, _) = Run(input, Streaming());
        Assert.Equal(FileStatus.Unchanged, result.Status);
        Assert.Contains(result.Residuals, r => r.Part == "xl/sharedStrings.xml");
    }

    [Fact]
    public void BinaryPartIsScannedInWindowsAsWhole()
    {
        var scanner = new ResidualScanner(Rules());
        var random = new Random(7);
        var bytes = new List<byte>();
        for (int i = 0; i < 60; i++)
        {
            var filler = new byte[random.Next(0, 300)];
            random.NextBytes(filler);
            bytes.AddRange(filler);
            string link = $"{Host}/Document/Fin/r{i}.xlsx";
            bytes.AddRange((i % 3) switch
            {
                0 => Encoding.Unicode.GetBytes(link),
                1 => Encoding.Latin1.GetBytes(link),
                _ => Encoding.Unicode.GetBytes(link.Replace("/", "\\")),
            });
        }
        byte[] data = [.. bytes];

        var whole = scanner.ScanBinary("xl/worksheets/sheet1.bin", data);
        foreach (int chunk in new[] { 2, 64, 1000, 4097 })
        {
            using var stream = new MemoryStream(data);
            var windows = scanner.ScanBinary("xl/worksheets/sheet1.bin", stream, chunk);
            Assert.Equal(whole.Select(r => (r.Encoding, r.Match, r.Context)).Order(), windows.Select(r => (r.Encoding, r.Match, r.Context)).Order());
        }
        Assert.True(whole.Count() >= 60);
    }

    [Theory]
    [InlineData("<c r=\"A1\"><v>1</v></c><c>", 22)]
    [InlineData("<x:c r=\"A1\"><x:v>1</x:v></x:c><x", 30)]
    [InlineData("<c r=\"A1\"><f>\"abc>\"</f>", -1)]
    [InlineData("</row><c>", -1)]
    public void CellEndIsFound(string text, int expected) =>
        Assert.Equal(expected, Text.TextSegments.AfterLastCellEnd(text));
}
