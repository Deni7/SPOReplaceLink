using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Rules;
using SPOReplaceLink.OpenXml.Parts;
using SPOReplaceLink.OpenXml.Text;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.OpenXml;

/// <summary>
/// Replaces links in one xlsx/xlsm/xlsb file without Excel (P-2): finds links in the parts that hold them,
/// rewrites only the changed parts, copies the rest as is (OOX-10), checks the result (OOX-14)
/// and looks for old prefixes left anywhere in the file (RES-1).
/// </summary>
public sealed partial class WorkbookProcessor
{
    private readonly RuleSet _rules;
    private readonly PlacesOptions _places;
    private readonly ResidualScanner _scanner;
    private readonly long _largePart;
    private readonly int _segmentChars;

    /// <summary>Parts bigger than this (uncompressed) are read piece by piece instead of as one string.</summary>
    internal const long DefaultLargePart = 64L * 1024 * 1024;

    /// <summary>Size of one piece of a large XML part, in chars.</summary>
    internal const int DefaultSegmentChars = 8 * 1024 * 1024;

    public WorkbookProcessor(RuleSet rules, PlacesOptions places)
        : this(rules, places, DefaultLargePart, DefaultSegmentChars)
    {
    }

    internal WorkbookProcessor(RuleSet rules, PlacesOptions places, long largePart, int segmentChars)
    {
        _rules = rules;
        _places = places;
        _scanner = new ResidualScanner(rules);
        _largePart = largePart;
        _segmentChars = segmentChars;
    }

    /// <param name="inputPath">The file to process; never modified.</param>
    /// <param name="outputPath">Where the modified copy goes if there are changes.</param>
    /// <param name="documentFolder">Folder the document lived in before migration, for relative links (RUL-6).</param>
    /// <param name="displayPath">Path shown in reports.</param>
    /// <param name="local">Set for an intermediate copy saved by Excel: how to undo the links Excel made relative to it.</param>
    public FileResult Process(string inputPath, string outputPath, string? documentFolder, string displayPath, LocalCopy? local = null)
    {
        string ext = Path.GetExtension(inputPath).ToLowerInvariant();
        var result = new FileResult { Path = displayPath, Format = ext.TrimStart('.') };
        Plan? plan = null;
        try
        {
            switch (PackageSniffer.Detect(inputPath))
            {
                case PackageKind.Compound when ext == ".xls" && PackageSniffer.DetectXlsPassword(inputPath) is var password && password != XlsPassword.None:
                    // Never handed to Excel; a retry run with the filter skipped:encrypted takes it up once the password is removed.
                    result.Status = FileStatus.SkippedEncrypted;
                    result.Error = password == XlsPassword.Open
                        ? L.T("Файл захищений паролем на відкриття.")
                        : L.T("Файл захищений паролем на зміну (Excel питає пароль або «лише для читання»).");
                    return result;
                case PackageKind.Compound when ext == ".xls":
                    result.Status = FileStatus.NeedsExcel;
                    result.Error = L.T("Формат xls обробляється через Excel (ExcelWorker).");
                    return result;
                case PackageKind.Compound when PackageSniffer.IsEncryptedOoxml(inputPath):
                    result.Status = FileStatus.SkippedEncrypted;
                    result.Error = L.T("Файл захищений паролем на відкриття.");
                    return result;
                case PackageKind.Compound:
                    return Fail(result, ErrorClass.File, L.F("Файл у форматі OLE (xls?), а не OOXML, хоча має розширення {0}.", ext));
                case PackageKind.Unknown:
                    return Fail(result, ErrorClass.File, L.T("Файл не є zip-архівом OOXML (пошкоджений або інший формат)."));
            }

            plan = new Plan(outputPath);
            Analyze(plan, inputPath, documentFolder, ext == ".xlsb", local);
            result.Links.AddRange(plan.Links);
            result.Residuals.AddRange(plan.Residuals);
            result.NeedsExcel = plan.NeedsExcel;
            if (!plan.HasChanges)
            {
                result.Status = FileStatus.Unchanged;
                return result;
            }

            result.ChangedParts.AddRange(plan.Changed.Keys);
            result.ChangedParts.AddRange(plan.Spilled.Keys);
            try
            {
                Write(inputPath, outputPath, plan);
                Verify(inputPath, outputPath, plan, documentFolder, local);
            }
            catch
            {
                TryDelete(outputPath);
                throw;
            }
            result.Status = FileStatus.Planned;
            result.OutputWritten = true;
            return result;
        }
        catch (VerificationException ex)
        {
            return Fail(result, ErrorClass.Validation, L.T("Самоперевірка не пройдена: ") + ex.Message);
        }
        catch (Exception ex) when (ex is InvalidDataException or DecoderFallbackException or EndOfStreamException)
        {
            return Fail(result, ErrorClass.File, L.T("Пошкоджений архів або частина: ") + ex.Message);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Fail(result, ErrorClass.File, ex.Message);
        }
        catch (Exception ex)
        {
            return Fail(result, ErrorClass.Internal, ex.GetType().Name + ": " + ex.Message);
        }
        finally
        {
            plan?.DeleteTemp();
        }
    }

    private sealed class Plan(string outputPath)
    {
        public Dictionary<string, byte[]> Changed { get; } = new(StringComparer.Ordinal);
        /// <summary>Changed large parts, written to temporary files next to the output instead of held in memory.</summary>
        public Dictionary<string, string> Spilled { get; } = new(StringComparer.Ordinal);
        public List<LinkRecord> Links { get; } = [];
        public List<Residual> Residuals { get; } = [];
        public bool NeedsExcel { get; set; }
        public bool HasChanges => Changed.Count > 0 || Spilled.Count > 0;
        private readonly List<string> _temp = [];

        public string NewTemp()
        {
            string path = outputPath + "." + _temp.Count + ".part";
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            _temp.Add(path);
            return path;
        }

        public void DeleteTemp()
        {
            foreach (string path in _temp)
                TryDelete(path);
        }
    }

    private enum PartKind { Rels, ExternalLink, Sheet, Workbook, Connections, CustomXml, OtherXml, Binary, Skip }

    [GeneratedRegex(@"^xl/externalLinks/externalLink\d+\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ExternalLinkPart();

    [GeneratedRegex(@"^xl/(worksheets|macrosheets|dialogsheets)/[^/]+\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex SheetPart();

    [GeneratedRegex(@"^customXml/item\d+\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex CustomXmlItem();

    private static PartKind KindOf(string name)
    {
        if (name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            return PartKind.Rels;
        if (ExternalLinkPart().IsMatch(name))
            return PartKind.ExternalLink;
        if (SheetPart().IsMatch(name))
            return PartKind.Sheet;
        if (name.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase))
            return PartKind.Workbook;
        if (name.Equals("xl/connections.xml", StringComparison.OrdinalIgnoreCase))
            return PartKind.Connections;
        if (CustomXmlItem().IsMatch(name))
            return PartKind.CustomXml;
        string ext = Path.GetExtension(name).ToLowerInvariant();
        return ext switch
        {
            ".xml" or ".vml" => PartKind.OtherXml,
            ".bin" => PartKind.Binary,
            _ => PartKind.Skip,
        };
    }

    private void Analyze(Plan plan, string path, string? documentFolder, bool isXlsb, LocalCopy? local)
    {
        (string Name, PartText Text)? workbook = null;
        using var zip = ZipFile.OpenRead(path);
        foreach (var entry in zip.Entries)
        {
            string name = entry.FullName;
            var kind = KindOf(name);
            if (kind == PartKind.Skip || name.EndsWith('/'))
                continue;

            if (kind is PartKind.Sheet or PartKind.OtherXml or PartKind.Binary && !IsWorkbookBin(name) && entry.Length > _largePart)
            {
                AnalyzeLarge(plan, entry, kind, documentFolder, isXlsb, local);
                continue;
            }

            byte[] bytes = ReadAll(entry);
            switch (kind)
            {
                case PartKind.Rels or PartKind.ExternalLink or PartKind.Sheet or PartKind.Connections:
                {
                    var text = PartText.Decode(bytes);
                    var ctx = new LinkContext(_rules, _places, documentFolder, name, local);
                    string updated = RewritePart(kind, text.Text, ctx);
                    plan.Links.AddRange(ctx.Links);
                    if (ctx.Changes > 0)
                        plan.Changed[name] = text.Encode(updated);
                    plan.Residuals.AddRange(_scanner.ScanText(name, updated));
                    break;
                }
                case PartKind.Workbook:
                {
                    var text = PartText.Decode(bytes);
                    workbook = (name, text);
                    plan.Residuals.AddRange(_scanner.ScanText(name, PartRewriters.MaskWorkbookPath(text.Text)));
                    break;
                }
                case PartKind.CustomXml:
                {
                    var text = PartText.Decode(bytes);
                    if (!PowerQuery.IsMashup(text.Text))
                    {
                        plan.Residuals.AddRange(_scanner.ScanText(name, text.Text));
                        break;
                    }
                    var ctx = new LinkContext(_rules, _places, documentFolder, name, local);
                    string updated = PowerQuery.Rewrite(text.Text, ctx);
                    plan.Links.AddRange(ctx.Links);
                    if (ctx.Changes > 0)
                        plan.Changed[name] = text.Encode(updated);
                    // The formulas are base64 in the part: old prefixes are looked for in the decoded texts.
                    foreach (var (inner, content) in PowerQuery.Contents(updated))
                        plan.Residuals.AddRange(_scanner.ScanText(name + "!" + inner, content));
                    break;
                }
                case PartKind.OtherXml:
                    plan.Residuals.AddRange(_scanner.ScanText(name, Encoding.UTF8.GetString(bytes)));
                    break;
                case PartKind.Binary:
                {
                    if (isXlsb && IsWorkbookBin(name))
                        bytes = XlsbRecords.MaskRecords(bytes, XlsbRecords.AbsPath15);
                    var found = _scanner.ScanBinary(name, bytes).ToList();
                    plan.Residuals.AddRange(found);
                    // XLB-2: old prefixes inside xlsb records can only be fixed by Excel.
                    if (isXlsb && found.Count > 0 && IsXlsbRecordPart(name))
                        plan.NeedsExcel = true;
                    break;
                }
            }
        }

        // The saved-at folder is refreshed only together with real changes: on its own it is not worth a new version.
        if (plan.HasChanges && workbook is { } wb)
        {
            var ctx = new LinkContext(_rules, _places, documentFolder, wb.Name, local);
            string updated = PartRewriters.RewriteWorkbookPath(wb.Text.Text, ctx);
            plan.Links.AddRange(ctx.Links);
            if (ctx.Changes > 0)
                plan.Changed[wb.Name] = wb.Text.Encode(updated);
        }
    }

    private static bool IsWorkbookBin(string name) => name.Equals("xl/workbook.bin", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// A part too big for one string (a sheet or sharedStrings of millions of rows, an xlsb sheet): read piece by piece.
    /// A changed sheet goes to a temporary file.
    /// </summary>
    private void AnalyzeLarge(Plan plan, ZipArchiveEntry entry, PartKind kind, string? documentFolder, bool isXlsb, LocalCopy? local)
    {
        string name = entry.FullName;
        if (kind == PartKind.Binary)
        {
            IReadOnlyList<Residual> found;
            using (var stream = entry.Open())
                found = _scanner.ScanBinary(name, stream);
            plan.Residuals.AddRange(found);
            if (isXlsb && found.Count > 0 && IsXlsbRecordPart(name))
                plan.NeedsExcel = true;
            return;
        }

        if (kind == PartKind.OtherXml)
        {
            using var reader = new StreamReader(entry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: false);
            foreach (string segment in TextSegments.Read(reader, _segmentChars, TextSegments.AfterLastTag))
                plan.Residuals.AddRange(_scanner.ScanText(name, segment));
            return;
        }

        var ctx = new LinkContext(_rules, _places, documentFolder, name, local);
        string temp = plan.NewTemp();
        using (var reader = OpenText(entry, out var writeEncoding))
        using (var writer = new StreamWriter(temp, append: false, writeEncoding, 1 << 16))
        {
            foreach (string segment in TextSegments.Read(reader, _segmentChars, TextSegments.AfterLastCellEnd))
            {
                string updated = PartRewriters.RewriteSheet(segment, ctx);
                writer.Write(updated);
                plan.Residuals.AddRange(_scanner.ScanText(name, updated));
            }
        }
        plan.Links.AddRange(ctx.Links);
        if (ctx.Changes > 0)
            plan.Spilled[name] = temp;
    }

    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    /// <summary>
    /// Text of a large XML part read the way <see cref="PartText.Decode"/> reads it, and the encoding (with its byte
    /// order mark, if the part had one) to write it back with.
    /// </summary>
    private static StreamReader OpenText(ZipArchiveEntry entry, out Encoding writeEncoding)
    {
        Span<byte> head = stackalloc byte[3];
        int read;
        using (var probe = entry.Open())
            read = probe.ReadAtLeast(head, 3, throwOnEndOfStream: false);
        Encoding readEncoding;
        int bom;
        if (read >= 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF)
            (readEncoding, writeEncoding, bom) = (StrictUtf8, new UTF8Encoding(true, true), 3);
        else if (read >= 2 && head[0] == 0xFF && head[1] == 0xFE)
            (readEncoding, writeEncoding, bom) = (new UnicodeEncoding(false, false), new UnicodeEncoding(false, true), 2);
        else if (read >= 2 && head[0] == 0xFE && head[1] == 0xFF)
            (readEncoding, writeEncoding, bom) = (new UnicodeEncoding(true, false), new UnicodeEncoding(true, true), 2);
        else
            (readEncoding, writeEncoding, bom) = (StrictUtf8, StrictUtf8, 0);

        var stream = entry.Open();
        try
        {
            stream.ReadExactly(new byte[bom]);
            return new StreamReader(stream, readEncoding, detectEncodingFromByteOrderMarks: false, 1 << 16);
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    /// <summary>
    /// XLS-3: links as Excel itself reports them for a written file (Edit Links, hyperlinks, formulas); every one that the
    /// rules would still change is a residual. This is how xls, which has no parts to read, is checked.
    /// </summary>
    public IReadOnlyList<Residual> CheckLinks(IEnumerable<(string Place, LinkKind Kind, string Value)> links, string? documentFolder, LocalCopy? local)
    {
        var residuals = new List<Residual>();
        foreach (var (place, kind, value) in links)
        {
            var ctx = new LinkContext(_rules, _places, documentFolder, place, local);
            if (kind == LinkKind.Formula)
                PartRewriters.RewriteFormula(Text.XmlText.EscapeText(value), ctx);
            else
                ctx.Rewrite(value, kind);
            foreach (var link in ctx.Links.Where(l => l.Action == LinkAction.Replaced))
                residuals.Add(new Residual(place, "excel", link.OldValue, value));
        }
        return residuals;
    }

    /// <summary>RES-1 for a file that is not a zip package (xls): old prefixes anywhere in its bytes.</summary>
    public IReadOnlyList<Residual> ScanFile(string path, string part)
    {
        using var stream = File.OpenRead(path);
        return _scanner.ScanBinary(part, stream);
    }

    private static bool IsXlsbRecordPart(string name) =>
        name.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
        && !name.StartsWith("xl/embeddings/", StringComparison.OrdinalIgnoreCase)
        && !name.Contains("vbaProject", StringComparison.OrdinalIgnoreCase);

    private static string RewritePart(PartKind kind, string text, LinkContext ctx) => kind switch
    {
        PartKind.Rels => PartRewriters.RewriteRels(text, ctx),
        PartKind.ExternalLink => PartRewriters.RewriteExternalLink(text, ctx),
        PartKind.Sheet => PartRewriters.RewriteSheet(text, ctx),
        PartKind.Workbook => PartRewriters.RewriteWorkbookPath(text, ctx),
        PartKind.Connections => PartRewriters.RewriteConnections(text, ctx),
        PartKind.CustomXml => PowerQuery.Rewrite(text, ctx),
        _ => text,
    };

    /// <summary>Copies the file and overwrites only the changed entries; the others keep their compressed bytes.</summary>
    private static void Write(string input, string output, Plan plan)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(output))!);
        File.Copy(input, output, overwrite: true);
        File.SetAttributes(output, FileAttributes.Normal);
        using var fs = new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Update);
        foreach (var (name, bytes) in plan.Changed)
        {
            using var stream = OpenForUpdate(zip, name);
            stream.Write(bytes);
        }
        foreach (var (name, temp) in plan.Spilled)
        {
            using var stream = OpenForUpdate(zip, name);
            using var source = File.OpenRead(temp);
            source.CopyTo(stream);
        }
    }

    private static Stream OpenForUpdate(ZipArchive zip, string name)
    {
        var entry = zip.GetEntry(name) ?? throw new InvalidOperationException(L.F("Частина {0} зникла з архіву.", name));
        var stream = entry.Open();
        stream.SetLength(0);
        return stream;
    }

    /// <summary>
    /// OOX-14: the written file opens, has the same parts, unchanged parts are byte-identical,
    /// changed parts are well-formed XML with exactly the expected links, and a second pass changes nothing.
    /// </summary>
    private void Verify(string input, string output, Plan plan, string? documentFolder, LocalCopy? local)
    {
        using var before = ZipFile.OpenRead(input);
        using var after = ZipFile.OpenRead(output);
        var a = before.Entries.ToList();
        var b = after.Entries.ToList();
        if (a.Count != b.Count || !a.Select(e => e.FullName).SequenceEqual(b.Select(e => e.FullName)))
            throw new VerificationException(L.T("склад частин архіву змінився."));

        for (int i = 0; i < a.Count; i++)
        {
            string name = b[i].FullName;
            if (plan.Spilled.TryGetValue(name, out string? temp))
            {
                VerifyLarge(b[i], temp, plan, documentFolder, local);
                continue;
            }
            if (!plan.Changed.TryGetValue(name, out byte[]? expected))
            {
                // .NET 8 rewrites an empty part as a stored entry, so Excel's two-byte deflate stream becomes zero bytes;
                // the content is the same.
                if (a[i].Crc32 != b[i].Crc32 || a[i].Length != b[i].Length
                    || (a[i].CompressedLength != b[i].CompressedLength && a[i].Length != 0))
                    throw new VerificationException(L.F("частина {0} змінилася, хоча замін у ній немає.", name));
                continue;
            }

            byte[] written = ReadAll(b[i]);
            if (!written.AsSpan().SequenceEqual(expected))
                throw new VerificationException(L.F("вміст {0} не збігається із записаним.", name));
            var text = PartText.Decode(written);
            EnsureWellFormed(name, text.Text);

            var ctx = new LinkContext(_rules, _places, documentFolder, name, local);
            RewritePart(KindOf(name), text.Text, ctx);
            if (ctx.Changes != 0)
                throw new VerificationException(L.F("повторна обробка {0} знайшла ще {1} замін.", name, ctx.Changes));
            var expectedLinks = plan.Links.Where(l => l.Part == name).Select(l => l.NewValue ?? l.OldValue);
            if (!ctx.Links.Select(l => l.OldValue).SequenceEqual(expectedLinks))
                throw new VerificationException(L.F("посилання в {0} після запису не відповідають очікуваним.", name));
        }
    }

    /// <summary>OOX-14 for a large changed sheet, read piece by piece.</summary>
    private void VerifyLarge(ZipArchiveEntry written, string temp, Plan plan, string? documentFolder, LocalCopy? local)
    {
        string name = written.FullName;
        using (var actual = written.Open())
        using (var expected = File.OpenRead(temp))
        {
            if (!SameContent(actual, expected))
                throw new VerificationException(L.F("вміст {0} не збігається із записаним.", name));
        }
        using (var stream = written.Open())
            EnsureWellFormed(name, XmlReader.Create(stream, XmlSettings));

        var ctx = new LinkContext(_rules, _places, documentFolder, name, local);
        using (var reader = OpenText(written, out _))
        {
            foreach (string segment in TextSegments.Read(reader, _segmentChars, TextSegments.AfterLastCellEnd))
                PartRewriters.RewriteSheet(segment, ctx);
        }
        if (ctx.Changes != 0)
            throw new VerificationException(L.F("повторна обробка {0} знайшла ще {1} замін.", name, ctx.Changes));
        var expectedLinks = plan.Links.Where(l => l.Part == name).Select(l => l.NewValue ?? l.OldValue);
        if (!ctx.Links.Select(l => l.OldValue).SequenceEqual(expectedLinks))
            throw new VerificationException(L.F("посилання в {0} після запису не відповідають очікуваним.", name));
    }

    private static bool SameContent(Stream a, Stream b)
    {
        var x = new byte[1 << 16];
        var y = new byte[1 << 16];
        while (true)
        {
            int n = a.ReadAtLeast(x, x.Length, throwOnEndOfStream: false);
            int m = b.ReadAtLeast(y, y.Length, throwOnEndOfStream: false);
            if (n != m || !x.AsSpan(0, n).SequenceEqual(y.AsSpan(0, m)))
                return false;
            if (n < x.Length)
                return true;
        }
    }

    private static readonly XmlReaderSettings XmlSettings = new() { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };

    private static void EnsureWellFormed(string name, string xml) =>
        EnsureWellFormed(name, XmlReader.Create(new StringReader(xml), XmlSettings));

    private static void EnsureWellFormed(string name, XmlReader reader)
    {
        try
        {
            using (reader)
            {
                while (reader.Read())
                {
                }
            }
        }
        catch (XmlException ex)
        {
            throw new VerificationException(L.F("{0} — невалідний XML: {1}", name, ex.Message));
        }
    }

    private static byte[] ReadAll(ZipArchiveEntry entry)
    {
        if (entry.Length > Array.MaxLength)
            throw new InvalidDataException(L.F("частина {0} завелика ({1} байт) для обробки цілою.", entry.FullName, entry.Length));
        using var stream = entry.Open();
        using var ms = new MemoryStream(entry.Length > 0 && entry.Length < int.MaxValue ? (int)entry.Length : 0);
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    private static FileResult Fail(FileResult result, ErrorClass errorClass, string message)
    {
        result.Status = FileStatus.Failed;
        result.ErrorClass = errorClass;
        result.Error = message;
        result.ChangedParts.Clear();
        return result;
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
        }
    }
}

public sealed class VerificationException(string message) : Exception(message);
