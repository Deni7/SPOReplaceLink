using System.Globalization;
using System.Text;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Reports;
using SPOReplaceLink.Core.Urls;
using SPOReplaceLink.State;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Engine;

/// <summary>Reports of a run from the state store: REP-1, REP-2, REP-3, RES-2, RTY-8 and the prefix histogram of inventory.</summary>
public static class ReportExporter
{
    public static readonly string[] Names = ["files", "links", "residuals", "attention", "prefixes", "sample", "summary", "pilot"];

    /// <summary>Reports that are text, not tables.</summary>
    public static bool IsText(string name) => name is "summary" or "pilot";

    public static string FileName(string name) => name + (IsText(name) ? ".txt" : ".csv");

    /// <summary>Folders below the root (site, share, drive) kept in a prefix of the prefixes report.</summary>
    public const int DefaultPrefixDepth = 2;
    public const int MaxPrefixDepth = 10;

    public static void ExportAll(StateStore store, long runId, string directory, int prefixDepth = DefaultPrefixDepth)
    {
        Directory.CreateDirectory(directory);
        foreach (string name in Names)
        {
            using var stream = File.Create(Path.Combine(directory, FileName(name)));
            Write(store, runId, name, stream, prefixDepth);
        }
    }

    public static void Write(StateStore store, long runId, string name, Stream stream, int prefixDepth = DefaultPrefixDepth)
    {
        using var w = new StreamWriter(stream, new UTF8Encoding(true), leaveOpen: true);
        if (IsText(name))
        {
            w.Write(TextReport(store, runId, name));
            return;
        }
        foreach (var row in Rows(store, runId, name, prefixDepth))
            Row(w, row);
    }

    /// <summary>GUI-10: the same report as an Excel workbook.</summary>
    public static void WriteXlsx(StateStore store, long runId, string name, Stream stream, int prefixDepth = DefaultPrefixDepth) =>
        OpenXml.SimpleXlsxWriter.Write(stream, name, IsText(name)
            ? TextReport(store, runId, name).Split('\n').Select(l => (IReadOnlyList<string?>)[l.TrimEnd('\r')])
            : Rows(store, runId, name, prefixDepth));

    /// <summary>Header and rows of a report; prefixDepth is used by the prefixes report only.</summary>
    public static IEnumerable<IReadOnlyList<string?>> Rows(StateStore store, long runId, string name, int prefixDepth = DefaultPrefixDepth)
    {
        switch (name)
        {
            case "files":
                return FileRows(store.AllFiles(runId));
            case "links":
                return Prepend(["path", "part", "kind", "action", "old_value", "new_value", "rule", "resolved_from"],
                    store.AllLinks(runId).Select(x => (IReadOnlyList<string?>)[x.Path, x.Link.Part, x.Link.Kind.ToString(), x.Link.Action.ToString(),
                        x.Link.OldValue, x.Link.NewValue, x.Link.RuleIndex?.ToString(CultureInfo.InvariantCulture), x.Link.ResolvedFrom]));
            case "residuals":
                return Prepend(["path", "part", "encoding", "match", "context"],
                    store.AllResiduals(runId).Select(x => (IReadOnlyList<string?>)[x.Path, x.Residual.Part, x.Residual.Encoding, x.Residual.Match, x.Residual.Context]));
            case "attention":
                return Prepend(["id", "path", "status", "error_class", "message", "attempts", "action"],
                    store.AllFiles(runId).Where(NeedsAttention).Select(f => (IReadOnlyList<string?>)[N(f.Id), f.Path ?? f.Name, f.StatusCode ?? f.Status,
                        f.ErrorClass, f.Error, N(f.Attempts), Recommend(f)]));
            case "sample":
                return PilotReport.Sample(store, runId);
            case "prefixes":
                return Prepend(["prefix", "links", "files", "replaced", "no_rule", "unresolved"],
                    Prefixes(store.AllLinks(runId), prefixDepth).Select(p => (IReadOnlyList<string?>)[p.Prefix, N(p.Links), N(p.Files), N(p.Replaced), N(p.NoRule), N(p.Unresolved)]));
            default:
                throw new ArgumentException(L.F("Невідомий звіт «{0}».", name), nameof(name));
        }
    }

    /// <summary>GUI-10: any filtered view of the files table.</summary>
    public static IEnumerable<IReadOnlyList<string?>> FileRows(IEnumerable<FileRow> files) =>
        Prepend(["id", "path", "format", "size", "status", "error_class", "error", "attempts", "replacements", "residuals",
                "version_before", "etag_after", "app", "comment"],
            files.Select(f => (IReadOnlyList<string?>)[N(f.Id), f.Path ?? f.Name, Path.GetExtension(f.Name).TrimStart('.'), N(f.Size), f.StatusCode ?? f.Status,
                f.ErrorClass, f.Error, N(f.Attempts), N(f.Replacements), N(f.Residuals), f.VersionBefore, f.ETagAfter, f.App, f.Comment]));

    public static void WriteCsv(Stream stream, IEnumerable<IReadOnlyList<string?>> rows)
    {
        using var w = new StreamWriter(stream, new UTF8Encoding(true), leaveOpen: true);
        foreach (var row in rows)
            Row(w, row);
    }

    private static IEnumerable<IReadOnlyList<string?>> Prepend(string?[] header, IEnumerable<IReadOnlyList<string?>> rows) =>
        rows.Prepend(header);
    public static bool NeedsAttention(FileRow f)
    {
        string code = f.StatusCode ?? f.Status;
        return code.StartsWith("failed:", StringComparison.Ordinal)
            || code.EndsWith(":residual", StringComparison.Ordinal)
            || code.Contains("needs_excel", StringComparison.Ordinal)
            || code == FileStatuses.ChangesPending;
    }

    /// <summary>RTY-8: what to do with the file next.</summary>
    public static string Recommend(FileRow f)
    {
        string code = f.StatusCode ?? f.Status;
        return f.ErrorClass switch
        {
            "transient" => L.T("retry: тимчасова помилка SharePoint або мережі"),
            "conflict" => L.T("retry: файл змінювали під час обробки"),
            "locked" => L.T("retry пізніше: файл відкритий або взятий на check-out; попросити власника закрити"),
            "auth" => L.T("перевірити права застосунку (Sites.Selected) на сайт, потім retry"),
            "file" => L.T("перевірити файл вручну (пошкоджений, зашифрований або іншого формату) або ignore"),
            "validation" => L.T("самоперевірка не пройдена: передати файл розробнику, файл не змінено"),
            "internal" => L.T("помилка інструменту: передати розробнику журнал, потім retry"),
            _ when code.Contains("needs_excel", StringComparison.Ordinal) => L.T("потрібна обробка через Excel (ExcelWorker)"),
            _ when code.EndsWith(":residual", StringComparison.Ordinal) => L.T("доповнити правила за residuals.csv, потім retry"),
            _ when code == FileStatuses.ChangesPending => L.T("запустити apply для цих файлів"),
            _ => string.Empty,
        };
    }

    public sealed record PrefixStat(string Prefix, long Links, long Files, long Replaced, long NoRule, long Unresolved);

    /// <summary>
    /// Histogram of link locations up to <paramref name="depth"/> folders below the root (host, server, drive),
    /// including those no rule covers (FEATURES 7, inventory).
    /// </summary>
    public static IReadOnlyList<PrefixStat> Prefixes(IEnumerable<(string Path, LinkRecord Link)> links, int depth = DefaultPrefixDepth)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(depth, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(depth, MaxPrefixDepth);
        var stats = new Dictionary<string, (long Links, HashSet<string> Files, long Replaced, long NoRule, long Unresolved)>(StringComparer.OrdinalIgnoreCase);
        foreach (var (path, l) in links)
        {
            if (l.Kind == LinkKind.WorkbookPath)
                continue;
            string prefix = PrefixOf(l.ResolvedFrom ?? l.OldValue, depth);
            if (!stats.TryGetValue(prefix, out var s))
                s = (0, new HashSet<string>(StringComparer.OrdinalIgnoreCase), 0, 0, 0);
            s.Links++;
            s.Files.Add(path);
            if (l.Action == LinkAction.Replaced || l.Action == LinkAction.PlaceDisabled)
                s.Replaced++;
            else if (l.Action == LinkAction.NoRule)
                s.NoRule++;
            else if (l.Action == LinkAction.Unresolved)
                s.Unresolved++;
            stats[prefix] = s;
        }
        return stats.Select(x => new PrefixStat(x.Key, x.Value.Links, x.Value.Files.Count, x.Value.Replaced, x.Value.NoRule, x.Value.Unresolved))
            .OrderByDescending(x => x.Links).ThenBy(x => x.Prefix, StringComparer.Ordinal).ToList();
    }

    /// <summary>The root (host, \\server, drive) and up to <paramref name="depth"/> segments below it.</summary>
    internal static string PrefixOf(string link, int depth = DefaultPrefixDepth)
    {
        var form = LinkClassifier.Classify(link);
        string text = UrlText.Decode(link);
        switch (form)
        {
            case LinkForm.Http:
            {
                int host = text.IndexOf("://", StringComparison.Ordinal) + 3;
                int end = text.IndexOfAny(['?', '#'], host);
                var segments = text[host..(end < 0 ? text.Length : end)].Split('/');
                return text[..host] + string.Join('/', segments.Take(depth + 1)).TrimEnd('/');
            }
            case LinkForm.Unc or LinkForm.Drive or LinkForm.File:
            {
                string path = text.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ? NormalizedLink.Normalize(text) : text.Replace('/', '\\');
                var segments = path.Split(['\\', '/'], StringSplitOptions.RemoveEmptyEntries);
                string root = form == LinkForm.Drive ? string.Empty : "\\\\";
                return root + string.Join('\\', segments.Take(depth + 1));
            }
            case LinkForm.ServerRelative:
                return "/" + string.Join('/', text.Split('/', StringSplitOptions.RemoveEmptyEntries).Take(depth));
            case LinkForm.Relative:
                return L.T("(відносні)");
            default:
                return "(" + form.ToString().ToLowerInvariant() + ")";
        }
    }

    private static string TextReport(StateStore store, long runId, string name) =>
        name == "pilot" ? PilotReport.Text(store, runId) : Summary(store, runId);

    /// <summary>REP-3.</summary>
    public static string Summary(StateStore store, long runId)
    {
        var run = store.GetRun(runId) ?? throw new ArgumentException(L.F("Прогін {0} не знайдено.", runId));
        var counts = store.CountByStatus(runId);
        var (files, replacements, residuals, bytes) = store.Totals(runId);
        var duration = (run.FinishedAt ?? DateTimeOffset.UtcNow) - (run.StartedAt ?? run.CreatedAt);
        var sb = new StringBuilder();
        sb.AppendLine(string.Format(CultureInfo.InvariantCulture, L.T("Прогін {0} ({1}), профіль {2}"), run.Id, run.Mode, run.ProfileName));
        if (run.ParentId is { } parent)
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, L.T("Батьківський прогін: {0}"), parent));
        sb.AppendLine(string.Format(CultureInfo.InvariantCulture, L.T("Статус: {0}{1}"), run.Status, run.StopReason is null ? string.Empty : " — " + run.StopReason));
        sb.AppendLine(string.Format(CultureInfo.InvariantCulture, L.T("Тривалість: {0:d\\.hh\\:mm\\:ss}"), duration));
        sb.AppendLine(string.Format(CultureInfo.InvariantCulture, L.T("Файлів: {0}, обсяг {1:0.0} МБ, замін посилань: {2}, залишків: {3}"), files, bytes / 1024.0 / 1024.0, replacements, residuals));
        if (duration.TotalMinutes > 0)
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, L.T("Швидкість: {0:0.0} файлів/хв"), files / Math.Max(duration.TotalMinutes, 1.0 / 60)));
        sb.AppendLine(L.T("За статусами:"));
        foreach (var (status, n) in counts.OrderBy(x => x.Key, StringComparer.Ordinal))
            sb.AppendLine(CultureInfo.InvariantCulture, $"  {status,-28} {n}");
        long attention = store.CountAttention(runId);
        sb.AppendLine(string.Format(CultureInfo.InvariantCulture, L.T("Потребують уваги: {0}"), attention));
        return sb.ToString();
    }

    private static string N(long n) => n.ToString(CultureInfo.InvariantCulture);

    private static void Row(StreamWriter w, IReadOnlyList<string?> values)
    {
        w.Write(string.Join(',', values.Select(v => CsvReport.Escape(v, ','))));
        w.Write("\r\n");
    }
}
