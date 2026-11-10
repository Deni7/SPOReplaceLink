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

    public WorkbookProcessor(RuleSet rules, PlacesOptions places)
    {
        _rules = rules;
        _places = places;
        _scanner = new ResidualScanner(rules);
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

            var plan = Analyze(inputPath, documentFolder, ext == ".xlsb", local);
            result.Links.AddRange(plan.Links);
            result.Residuals.AddRange(plan.Residuals);
            result.NeedsExcel = plan.NeedsExcel;
            if (plan.Changed.Count == 0)
            {
                result.Status = FileStatus.Unchanged;
                return result;
            }

            result.ChangedParts.AddRange(plan.Changed.Keys);
            try
            {
                Write(inputPath, outputPath, plan.Changed);
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
    }

    private sealed class Plan
    {
        public Dictionary<string, byte[]> Changed { get; } = new(StringComparer.Ordinal);
        public List<LinkRecord> Links { get; } = [];
        public List<Residual> Residuals { get; } = [];
        public bool NeedsExcel { get; set; }
    }

    private enum PartKind { Rels, ExternalLink, Sheet, Workbook, OtherXml, Binary, Skip }

    [GeneratedRegex(@"^xl/externalLinks/externalLink\d+\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ExternalLinkPart();

    [GeneratedRegex(@"^xl/(worksheets|macrosheets|dialogsheets)/[^/]+\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex SheetPart();

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
        string ext = Path.GetExtension(name).ToLowerInvariant();
        return ext switch
        {
            ".xml" or ".vml" => PartKind.OtherXml,
            ".bin" => PartKind.Binary,
            _ => PartKind.Skip,
        };
    }

    private Plan Analyze(string path, string? documentFolder, bool isXlsb, LocalCopy? local)
    {
        var plan = new Plan();
        (string Name, PartText Text)? workbook = null;
        using var zip = ZipFile.OpenRead(path);
        foreach (var entry in zip.Entries)
        {
            string name = entry.FullName;
            var kind = KindOf(name);
            if (kind == PartKind.Skip || name.EndsWith('/'))
                continue;

            byte[] bytes = ReadAll(entry);
            switch (kind)
            {
                case PartKind.Rels or PartKind.ExternalLink or PartKind.Sheet:
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
                case PartKind.OtherXml:
                    plan.Residuals.AddRange(_scanner.ScanText(name, Encoding.UTF8.GetString(bytes)));
                    break;
                case PartKind.Binary:
                {
                    if (isXlsb && name.Equals("xl/workbook.bin", StringComparison.OrdinalIgnoreCase))
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
        if (plan.Changed.Count > 0 && workbook is { } wb)
        {
            var ctx = new LinkContext(_rules, _places, documentFolder, wb.Name, local);
            string updated = PartRewriters.RewriteWorkbookPath(wb.Text.Text, ctx);
            plan.Links.AddRange(ctx.Links);
            if (ctx.Changes > 0)
                plan.Changed[wb.Name] = wb.Text.Encode(updated);
        }
        return plan;
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
    public IReadOnlyList<Residual> ScanFile(string path, string part) => _scanner.ScanBinary(part, File.ReadAllBytes(path)).ToList();

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
        _ => text,
    };

    /// <summary>Copies the file and overwrites only the changed entries; the others keep their compressed bytes.</summary>
    private static void Write(string input, string output, Dictionary<string, byte[]> changed)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(output))!);
        File.Copy(input, output, overwrite: true);
        File.SetAttributes(output, FileAttributes.Normal);
        using var fs = new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Update);
        foreach (var (name, bytes) in changed)
        {
            var entry = zip.GetEntry(name) ?? throw new InvalidOperationException(L.F("Частина {0} зникла з архіву.", name));
            using var stream = entry.Open();
            stream.SetLength(0);
            stream.Write(bytes);
        }
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
            byte[] written = ReadAll(b[i]);
            if (!plan.Changed.TryGetValue(name, out byte[]? expected))
            {
                // .NET 8 rewrites an empty part as a stored entry, so Excel's two-byte deflate stream becomes zero bytes;
                // the content is the same.
                if (a[i].Crc32 != b[i].Crc32 || a[i].Length != b[i].Length
                    || (a[i].CompressedLength != b[i].CompressedLength && a[i].Length != 0))
                    throw new VerificationException(L.F("частина {0} змінилася, хоча замін у ній немає.", name));
                continue;
            }

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

    private static void EnsureWellFormed(string name, string xml)
    {
        try
        {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            using var reader = XmlReader.Create(new StringReader(xml), settings);
            while (reader.Read())
            {
            }
        }
        catch (XmlException ex)
        {
            throw new VerificationException(L.F("{0} — невалідний XML: {1}", name, ex.Message));
        }
    }

    private static byte[] ReadAll(ZipArchiveEntry entry)
    {
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
