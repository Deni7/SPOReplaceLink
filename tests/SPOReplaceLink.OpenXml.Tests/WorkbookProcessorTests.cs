using System.IO.Compression;
using System.Text;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Rules;
using static SPOReplaceLink.OpenXml.Tests.TestWorkbook;

namespace SPOReplaceLink.OpenXml.Tests;

public sealed class WorkbookProcessorTests : IDisposable
{
    private const string OldFolder = Host + "/Document/Fin";
    private const string SheetRels = "xl/worksheets/_rels/sheet1.xml.rels";
    private const string Sheet = "xl/worksheets/sheet1.xml";
    private const string ExtRels = "xl/externalLinks/_rels/externalLink1.xml.rels";

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
            new("/sites/finance/Analytics/", "/sites/finance/analytics_migrated/"),
        ], Host);
        Assert.True(build.IsValid, string.Join("\n", build.Issues));
        return build.RuleSet!;
    }

    private (FileResult Result, string Output) Run(TestWorkbook wb, string ext = ".xlsx", string? folder = OldFolder, PlacesOptions? places = null)
    {
        string input = wb.Save(Path.Combine(_dir, "in", Guid.NewGuid().ToString("N") + ext));
        return Run(input, folder, places);
    }

    private (FileResult Result, string Output) Run(string input, string? folder = OldFolder, PlacesOptions? places = null)
    {
        string output = Path.Combine(_dir, "out", Guid.NewGuid().ToString("N") + Path.GetExtension(input));
        var result = new WorkbookProcessor(Rules(), places ?? new PlacesOptions()).Process(input, output, folder, Path.GetFileName(input));
        Assert.Equal(result.OutputWritten, File.Exists(output));
        return (result, output);
    }

    [Fact]
    public void RewritesLinksInAllPlaces()
    {
        var (result, output) = Run(Default());

        Assert.Equal(FileStatus.Planned, result.Status);
        Assert.Equal("planned", result.StatusCode);
        Assert.Equal(2, result.ReplacementsOf(LinkKind.Hyperlink));
        Assert.Equal(1, result.ReplacementsOf(LinkKind.LinkSource));
        Assert.Equal(2, result.ReplacementsOf(LinkKind.Formula));
        Assert.Equal(1, result.ReplacementsOf(LinkKind.CachedValue));
        Assert.Equal([ExtRels, SheetRels, Sheet], result.ChangedParts.Order(StringComparer.Ordinal));

        string rels = ReadPart(output, SheetRels);
        Assert.Contains($"Target=\"{Host}/sites/finance/Fin/a%20b.xlsx?web=1&amp;x=2\"", rels);
        Assert.Contains($"Target=\"{Host}/sites/finance/analytics_migrated/rel.xlsx\"", rels);
        Assert.Contains($"Target=\"{Host}/sites/finance/analytics_migrated/Source.xlsx\"", ReadPart(output, ExtRels));

        string sheet = ReadPart(output, Sheet);
        // A raw-spelled link stays raw (RUL-5).
        Assert.Contains($"HYPERLINK(\"{Host}/sites/finance/Fin/Звіт 1.xlsx\",\"Звіт\")", sheet);
        Assert.Contains($"<f>\"{Host}/sites/finance/analytics_migrated/\"&amp;\"x.xlsx\"</f>", sheet);
        Assert.Contains($"<v>{Host}/sites/finance/analytics_migrated/x.xlsx</v>", sheet);
        Assert.Contains("<f>[1]Sheet1!A1</f>", sheet);
    }

    [Fact]
    public void ChangedPartsDifferOnlyInReplacedValues()
    {
        var wb = Default();
        var (_, output) = Run(wb);

        string expectedRels = wb.Text(SheetRels)
            .Replace($"{Host}/Document/Fin/a%20b.xlsx", $"{Host}/sites/finance/Fin/a%20b.xlsx")
            .Replace("../Analytics/rel.xlsx", $"{Host}/sites/finance/analytics_migrated/rel.xlsx");
        Assert.Equal(expectedRels, ReadPart(output, SheetRels));

        string expectedSheet = wb.Text(Sheet)
            .Replace($"{Host}/Document/Fin/", $"{Host}/sites/finance/Fin/")
            .Replace($"{Host}/Document/Analytics/", $"{Host}/sites/finance/analytics_migrated/");
        Assert.Equal(expectedSheet, ReadPart(output, Sheet));
    }

    [Fact]
    public void UnchangedPartsAreCopiedAsIs()
    {
        var wb = Default().Set("xl/vbaProject.bin", Enumerable.Range(0, 5000).Select(i => (byte)(i * 7)).ToArray());
        string input = wb.Save(Path.Combine(_dir, "in", "macro.xlsm"));
        var (result, output) = Run(input);
        Assert.True(result.OutputWritten);

        using var before = ZipFile.OpenRead(input);
        using var after = ZipFile.OpenRead(output);
        Assert.Equal(before.Entries.Select(e => e.FullName), after.Entries.Select(e => e.FullName));
        foreach (var a in before.Entries.Where(e => !result.ChangedParts.Contains(e.FullName)))
        {
            var b = after.GetEntry(a.FullName)!;
            Assert.Equal(a.Crc32, b.Crc32);
            Assert.Equal(a.CompressedLength, b.CompressedLength);
            Assert.Equal(a.LastWriteTime, b.LastWriteTime);
        }
        Assert.Equal(wb.Parts["xl/vbaProject.bin"], ReadParts(output)["xl/vbaProject.bin"]);
    }

    [Fact]
    public void EmptyPartDeflatedByExcelPassesTheSelfCheck()
    {
        const string Empty = "xl/customProperty3.bin";
        string input = Default().Set(Empty, []).SaveLikeExcel(Path.Combine(_dir, "in", "custom.xlsx"));
        using (var zip = ZipFile.OpenRead(input))
            Assert.Equal(2, zip.GetEntry(Empty)!.CompressedLength);

        var (result, output) = Run(input);
        Assert.Equal(FileStatus.Planned, result.Status);
        Assert.Empty(ReadParts(output)[Empty]);
    }

    [Fact]
    public void SecondRunChangesNothing()
    {
        var (_, output) = Run(Default());
        var (again, againOutput) = Run(output);
        Assert.Equal(FileStatus.Unchanged, again.Status);
        Assert.False(again.OutputWritten);
        Assert.False(File.Exists(againOutput));
        Assert.Empty(again.Residuals);
    }

    [Fact]
    public void ReportsResidualsInPartsThatAreNotRewritten()
    {
        var wb = Default().Set("xl/sharedStrings.xml",
            $"<sst xmlns=\"{NsMain}\"><si><t>Див. {Host}/Document/Fin/old.xlsx тут</t></si><si><t>/Document/Звіт</t></si></sst>");
        var (result, _) = Run(wb);

        Assert.Equal("planned:residual", result.StatusCode);
        Assert.Equal(2, result.Residuals.Count);
        Assert.All(result.Residuals, r => Assert.Equal("xl/sharedStrings.xml", r.Part));
        Assert.Contains(result.Residuals, r => r.Match == Host + "/Document" && r.Context.Contains("Див."));
        Assert.True(result.NeedsAttention);
    }

    [Fact]
    public void ResidualMatchesEncodedAndBackslashSpellings()
    {
        var wb = Default().Set("xl/connections.xml",
            "<connections><c a=\"HTTPS:%2F%2Fcontoso.sharepoint.com%2FDocument%2Fx\"/><c b=\"\\Document\\y\"/><c d=\"/Documents/z\"/><c e=\"/sites/finance/Document/q\"/></connections>");
        var (result, _) = Run(wb);
        Assert.Equal(2, result.Residuals.Count);
    }

    [Fact]
    public void DisabledPlaceIsLeftAndReported()
    {
        var (result, output) = Run(Default(), places: new PlacesOptions { Formulas = false });

        Assert.Equal(0, result.ReplacementsOf(LinkKind.Formula));
        Assert.Contains(result.Links, l => l.Kind == LinkKind.Formula && l.Action == LinkAction.PlaceDisabled);
        Assert.DoesNotContain(Sheet, result.ChangedParts);
        Assert.Contains(result.Residuals, r => r.Part == Sheet);
        Assert.Contains($"{Host}/Document/Fin/Звіт 1.xlsx", ReadPart(output, Sheet));
    }

    [Fact]
    public void RelativeLinkWithoutOldLocationIsLeft()
    {
        var (result, output) = Run(Default(), folder: null);
        Assert.Contains(result.Links, l => l.OldValue == "../Analytics/rel.xlsx" && l.Action == LinkAction.Unresolved);
        Assert.Contains("Target=\"../Analytics/rel.xlsx\"", ReadPart(output, SheetRels));
    }

    [Fact]
    public void InternalRelationshipsAreNotTouched()
    {
        var wb = Default().Set("xl/_rels/workbook.xml.rels",
            wb0().Text("xl/_rels/workbook.xml.rels").Replace("Target=\"sharedStrings.xml\"", "Target=\"/Document/sharedStrings.xml\""));
        var (result, _) = Run(wb);
        Assert.DoesNotContain("xl/_rels/workbook.xml.rels", result.ChangedParts);

        static TestWorkbook wb0() => Default();
    }

    [Fact]
    public void KeepsBomAndSingleQuotes()
    {
        string rels = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id='rId1' Type='{NsRel}/hyperlink' Target='{Host}/Document/it&apos;s.xlsx' TargetMode='External'/></Relationships>";
        var wb = Default().Set(SheetRels, [0xEF, 0xBB, 0xBF, .. Encoding.UTF8.GetBytes(rels)]);
        var (_, output) = Run(wb);

        byte[] bytes = ReadParts(output)[SheetRels];
        Assert.Equal([0xEF, 0xBB, 0xBF], bytes[..3]);
        Assert.Contains($"Target='{Host}/sites/finance/it&apos;s.xlsx'", Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3));
    }

    [Fact]
    public void HandlesPrefixedFormulaElementsAndEscapedQuotes()
    {
        string sheet = $"<x:worksheet xmlns:x=\"{NsMain}\"><x:sheetData><x:row r=\"1\"><x:c r=\"A1\" t=\"str\">" +
            $"<x:f>IF(A2=\"\"\"\",HYPERLINK(\"{Host}/Document/a\"\"b.xlsx\"),'It''s \"{Host}/Document\"'!A1)</x:f><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>";
        var (_, output) = Run(Default().Set(Sheet, sheet));
        string result = ReadPart(output, Sheet);
        Assert.Contains($"HYPERLINK(\"{Host}/sites/finance/a\"\"b.xlsx\")", result);
        // Text inside a quoted sheet name is not a string literal.
        Assert.Contains($"'It''s \"{Host}/Document\"'!A1", result);
    }

    [Fact]
    public void RewritesLinksEmbeddedInText()
    {
        string sheet = $"<worksheet xmlns=\"{NsMain}\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"str\">" +
            $"<f>\"Джерело: {Host}/Document/a.xlsx, оновлено\"</f></c></row></sheetData></worksheet>";
        var (_, output) = Run(Default().Set(Sheet, sheet));
        Assert.Contains($"\"Джерело: {Host}/sites/finance/a.xlsx, оновлено\"", ReadPart(output, Sheet));
    }

    [Fact]
    public void XlsbRelsAreRewrittenAndBinaryResidualsNeedExcel()
    {
        var wb = new TestWorkbook()
            .Set("[Content_Types].xml", Default().Text("[Content_Types].xml"))
            .Set("xl/worksheets/_rels/sheet1.bin.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                $"<Relationship Id=\"rId1\" Type=\"{NsRel}/hyperlink\" Target=\"{Host}/Document/a.xlsx\" TargetMode=\"External\"/></Relationships>")
            .Set("xl/worksheets/sheet1.bin", [0x01, 0x02, .. Encoding.Unicode.GetBytes($"{Host}/Document/b.xlsx"), 0x00]);
        var (result, output) = Run(wb, ".xlsb");

        Assert.Equal(1, result.ReplacementsOf(LinkKind.Hyperlink));
        Assert.True(result.NeedsExcel);
        Assert.Equal("planned:needs_excel", result.StatusCode);
        Assert.Contains(result.Residuals, r => r.Part == "xl/worksheets/sheet1.bin" && r.Encoding == "utf-16le");
        Assert.Contains($"{Host}/sites/finance/a.xlsx", ReadPart(output, "xl/worksheets/_rels/sheet1.bin.rels"));
    }

    [Fact]
    public void XlsbWithoutBinaryResidualsIsComplete()
    {
        var wb = new TestWorkbook()
            .Set("xl/worksheets/_rels/sheet1.bin.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                $"<Relationship Id=\"rId1\" Type=\"{NsRel}/hyperlink\" Target=\"{Host}/Document/a.xlsx\" TargetMode=\"External\"/></Relationships>")
            .Set("xl/worksheets/sheet1.bin", [0x01, 0x02, 0x03]);
        var (result, _) = Run(wb, ".xlsb");
        Assert.Equal("planned", result.StatusCode);
    }

    [Fact]
    public void EncryptedFileIsSkipped()
    {
        byte[] cfb = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, .. new byte[100],
            .. Encoding.Unicode.GetBytes("EncryptionInfo"), 0, 0, .. Encoding.Unicode.GetBytes("EncryptedPackage")];
        string input = Path.Combine(_dir, "enc.xlsx");
        Directory.CreateDirectory(_dir);
        File.WriteAllBytes(input, cfb);
        var (result, _) = Run(input);
        Assert.Equal("skipped:encrypted", result.StatusCode);
    }

    [Fact]
    public void CompoundFilesAreRoutedByExtension()
    {
        Directory.CreateDirectory(_dir);
        byte[] cfb = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, .. new byte[100]];
        File.WriteAllBytes(Path.Combine(_dir, "old.xls"), cfb);
        File.WriteAllBytes(Path.Combine(_dir, "renamed.xlsx"), cfb);

        Assert.Equal("needs_excel", Run(Path.Combine(_dir, "old.xls")).Result.StatusCode);
        Assert.Equal("failed:file", Run(Path.Combine(_dir, "renamed.xlsx")).Result.StatusCode);
    }

    /// <summary>Workbooks saved by Excel (tests\fixtures, password "secret"): a password is found without Excel.</summary>
    [Theory]
    [InlineData("plain.xls", XlsPassword.None, "needs_excel")]
    [InlineData("password-open.xls", XlsPassword.Open, "skipped:encrypted")]
    [InlineData("password-open-excel95.xls", XlsPassword.Open, "skipped:encrypted")]
    [InlineData("password-write.xls", XlsPassword.Modify, "skipped:encrypted")]
    public void XlsWithAPasswordIsSkippedBeforeExcel(string name, XlsPassword expected, string status)
    {
        string path = Path.Combine(AppContext.BaseDirectory, "fixtures", name);
        Assert.Equal(expected, PackageSniffer.DetectXlsPassword(path));
        Assert.Equal(status, Run(path).Result.StatusCode);
    }

    [Theory]
    [InlineData(true, XlsPassword.Open)]
    [InlineData(false, XlsPassword.None)]
    public void SmallWorkbookStreamIsReadFromTheMiniStream(bool encrypted, XlsPassword expected)
    {
        // BOF, then FILEPASS or another record; 100 bytes, so two 64-byte mini sectors.
        byte[] biff = [0x09, 0x08, 16, 0, .. new byte[16], (byte)(encrypted ? 0x2F : 0x5C), 0, 76, 0, .. new byte[76]];
        Directory.CreateDirectory(_dir);
        string path = Path.Combine(_dir, "small.xls");
        File.WriteAllBytes(path, MiniCompoundFile("Workbook", biff));
        Assert.Equal(expected, PackageSniffer.DetectXlsPassword(path));
    }

    [Fact]
    public void DamagedCompoundFileIsLeftToExcel()
    {
        Directory.CreateDirectory(_dir);
        string path = Path.Combine(_dir, "damaged.xls");
        byte[] cfb = MiniCompoundFile("Workbook", [0x09, 0x08, 0, 0, 0x2F, 0, 0, 0]);
        cfb[0x30] = 0x77; // directory in a sector that does not exist
        File.WriteAllBytes(path, cfb);
        Assert.Equal(XlsPassword.None, PackageSniffer.DetectXlsPassword(path));
        Assert.Equal("needs_excel", Run(path).Result.StatusCode);
    }

    /// <summary>A version 3 compound file with one stream kept in the mini stream: FAT, directory, mini FAT, mini stream.</summary>
    private static byte[] MiniCompoundFile(string name, byte[] content)
    {
        const uint end = 0xFFFFFFFE, free = 0xFFFFFFFF;
        var file = new byte[512 * 5];

        void U16(int at, int v) => System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(at), (ushort)v);
        void U32(int at, uint v) => System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(at), v);
        byte[] signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        signature.CopyTo(file, 0);
        U16(0x18, 0x3E); U16(0x1A, 3); U16(0x1C, 0xFFFE); U16(0x1E, 9); U16(0x20, 6);
        U32(0x2C, 1); U32(0x30, 1); U32(0x38, 4096); U32(0x3C, 2); U32(0x40, 1); U32(0x44, end);
        for (int i = 0; i < 109; i++)
            U32(0x4C + i * 4, i == 0 ? 0 : free);
        // Sector 0: FAT; 1: directory; 2: mini FAT; 3: mini stream.
        int fat = 512;
        for (int i = 0; i < 128; i++)
            U32(fat + i * 4, i switch { 0 => 0xFFFFFFFD, 1 or 2 or 3 => end, _ => free });
        int dir = 512 * 2;
        void Entry(int index, string entryName, byte type, uint start, uint size)
        {
            int at = dir + index * 128;
            Encoding.Unicode.GetBytes(entryName).CopyTo(file, at);
            U16(at + 0x40, (entryName.Length + 1) * 2);
            file[at + 0x42] = type;
            U32(at + 0x44, free); U32(at + 0x48, free); U32(at + 0x4C, free);
            U32(at + 0x74, start);
            U32(at + 0x78, size);
        }
        int sectors = (content.Length + 63) / 64;
        Entry(0, "Root Entry", 5, 3, (uint)(sectors * 64));
        Entry(1, name, 2, 0, (uint)content.Length);
        U32(dir + 0x4C, 1); // root's child
        int miniFat = 512 * 3;
        for (int i = 0; i < 128; i++)
            U32(miniFat + i * 4, i < sectors - 1 ? (uint)(i + 1) : i == sectors - 1 ? end : free);
        content.CopyTo(file, 512 * 4);
        return file;
    }

    [Theory]
    [InlineData("not a zip at all")]
    [InlineData("PK\u0003\u0004 truncated")]
    public void DamagedFilesFailWithFileClass(string content)
    {
        Directory.CreateDirectory(_dir);
        string input = Path.Combine(_dir, "bad.xlsx");
        File.WriteAllText(input, content);
        var (result, _) = Run(input);
        Assert.Equal("failed:file", result.StatusCode);
        Assert.NotNull(result.Error);
    }

    [Fact]
    public void MalformedXmlAfterReplacementFailsValidation()
    {
        // Unclosed element: the rewrite itself works on text, the self-check catches the broken part.
        string rels = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id=\"rId1\" Type=\"{NsRel}/hyperlink\" Target=\"{Host}/Document/a.xlsx\" TargetMode=\"External\"/>";
        var (result, _) = Run(Default().Set(SheetRels, rels));
        Assert.Equal("failed:validation", result.StatusCode);
        Assert.False(result.OutputWritten);
    }

    private static TestWorkbook WithAbsPath(TestWorkbook wb) =>
        wb.Set("xl/workbook.xml", wb.Text("xl/workbook.xml").Replace("<sheets>",
            "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x15\">" +
            $"<x15ac:absPath url=\"{Host}/Document/Fin/\" xmlns:x15ac=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac\"/>" +
            "</mc:Choice></mc:AlternateContent><sheets>"));

    [Fact]
    public void WorkbookPathIsUpdatedTogetherWithRealChanges()
    {
        var (result, output) = Run(WithAbsPath(Default()));
        Assert.Contains("xl/workbook.xml", result.ChangedParts);
        Assert.Contains($"<x15ac:absPath url=\"{Host}/sites/finance/Fin/\"", ReadPart(output, "xl/workbook.xml"));
        Assert.Equal(6, result.Replacements);
        Assert.Empty(result.Residuals);
    }

    [Fact]
    public void WorkbookPathAloneNeitherChangesFileNorCountsAsResidual()
    {
        var wb = WithAbsPath(Default())
            .Set(SheetRels, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>")
            .Set(Sheet, $"<worksheet xmlns=\"{NsMain}\"><sheetData/></worksheet>")
            .Set(ExtRels, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        var (result, _) = Run(wb);
        Assert.Equal("unchanged", result.StatusCode);
        Assert.Empty(result.Residuals);
    }

    [Fact]
    public void XlsbWorkbookPathRecordIsNotAResidual()
    {
        string path = $"{Host}/Document/Fin/";
        byte[] str = Encoding.Unicode.GetBytes(path);
        int size = 4 + str.Length;
        byte[] absPath = [0x97, 0x10, (byte)(size & 0x7F | 0x80), (byte)(size >> 7), .. BitConverter.GetBytes(path.Length), .. str];
        byte[] other = [0x01, 0x02, 0xAA, 0xBB];
        var wb = new TestWorkbook().Set("xl/workbook.bin", [.. other, .. absPath, .. other]);
        var (result, _) = Run(wb, ".xlsb");
        Assert.Equal("unchanged", result.StatusCode);

        wb.Set("xl/workbook.bin", [.. other, .. absPath, 0x05, (byte)(str.Length), .. str]);
        Assert.Equal("unchanged:needs_excel", Run(wb, ".xlsb").Result.StatusCode);
    }

    [Fact]
    public void NoChangesMeansNoOutput()
    {
        var wb = Default()
            .Set(SheetRels, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>")
            .Set(Sheet, $"<worksheet xmlns=\"{NsMain}\"><sheetData/></worksheet>")
            .Set(ExtRels, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        var (result, _) = Run(wb);
        Assert.Equal("unchanged", result.StatusCode);
        Assert.False(result.OutputWritten);
    }
}
