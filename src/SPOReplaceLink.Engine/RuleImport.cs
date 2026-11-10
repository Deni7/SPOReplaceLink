using System.Text;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Reports;
using SPOReplaceLink.OpenXml;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Engine;

/// <summary>
/// GUI-23: rules from a mapping table "old path → new path" (CSV or XLSX) or from a list exported from ReplaceMagic.
/// The exact ReplaceMagic export format is not known yet (open question 9), so text is read tolerantly:
/// two columns separated by tab, semicolon, comma or "|", or one line "old => new" / "old -> new" / "old → new".
/// A header row (old/new, from/to, старий/новий…) is skipped.
/// </summary>
public static class RuleImport
{
    public sealed record Result(List<RuleEntry> Rules, List<string> Warnings);

    private static readonly string[] Arrows = ["=>", "->", "→", "\t"];
    private static readonly string[] HeaderWords = ["old", "new", "from", "to", "find", "replace", "source", "target", "стар", "нов", "знайти", "замін"]; // i18n-ignore: header words matched in imported files

    public static Result Parse(string fileName, Stream content)
    {
        string ext = Path.GetExtension(fileName).ToLowerInvariant();
        List<string[]> rows;
        if (ext is ".xlsx" or ".xlsm")
        {
            rows = SimpleXlsxReader.ReadFirstSheet(content);
        }
        else
        {
            using var reader = new StreamReader(content, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            rows = reader.ReadToEnd().Split('\n').Select(l => SplitLine(l.TrimEnd('\r'))).ToList();
        }
        return FromRows(rows);
    }

    internal static Result FromRows(IEnumerable<string[]> rows)
    {
        var rules = new List<RuleEntry>();
        var warnings = new List<string>();
        int line = 0;
        foreach (var raw in rows)
        {
            line++;
            var cells = raw.Select(c => c.Trim().Trim('"').Trim()).ToArray();
            if (cells.All(c => c.Length == 0))
                continue;
            if (line == 1 && LooksLikeHeader(cells))
                continue;
            var values = cells.Where(c => c.Length > 0).ToArray();
            if (values.Length < 2)
            {
                warnings.Add(L.F("Рядок {0}: потрібні два значення (старий і новий префікс) — пропущено.", line));
                continue;
            }
            rules.Add(new RuleEntry { Old = values[0], New = values[1], Name = values.Length > 2 ? values[2] : null });
        }
        if (rules.Count == 0)
            warnings.Add(L.T("Не знайдено жодного правила."));
        return new Result(rules, warnings);
    }

    private static bool LooksLikeHeader(string[] cells) =>
        cells.Length >= 2 && !cells[0].Contains("://") && !cells[0].StartsWith('/') && !cells[0].StartsWith('\\')
        && HeaderWords.Any(w => cells[0].Contains(w, StringComparison.OrdinalIgnoreCase));

    internal static string[] SplitLine(string line)
    {
        foreach (string arrow in Arrows)
        {
            int i = line.IndexOf(arrow, StringComparison.Ordinal);
            if (i > 0)
                return [line[..i], .. SplitLine(line[(i + arrow.Length)..])];
        }
        // ';' and '|' never appear in URLs of SharePoint paths in practice; ',' only if nothing else separates.
        foreach (char sep in new[] { ';', '|', ',' })
        {
            if (line.Contains(sep))
                return SplitCsv(line, sep);
        }
        return [line];
    }

    private static string[] SplitCsv(string line, char sep)
    {
        var cells = new List<string>();
        var sb = new StringBuilder();
        bool quoted = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"')
            {
                if (quoted && i + 1 < line.Length && line[i + 1] == '"')
                {
                    sb.Append('"');
                    i++;
                }
                else
                {
                    quoted = !quoted;
                }
            }
            else if (c == sep && !quoted)
            {
                cells.Add(sb.ToString());
                sb.Clear();
            }
            else
            {
                sb.Append(c);
            }
        }
        cells.Add(sb.ToString());
        return [.. cells];
    }

    /// <summary>GUI-23: template rules as CSV (old,new,name), UTF-8 with BOM.</summary>
    public static byte[] ToCsv(IEnumerable<RuleEntry> rules)
    {
        var sb = new StringBuilder("old,new,name\r\n");
        foreach (var r in rules)
            sb.Append(CsvReport.Escape(r.Old, ',')).Append(',').Append(CsvReport.Escape(r.New, ',')).Append(',').Append(CsvReport.Escape(r.Name, ',')).Append("\r\n");
        return [.. Encoding.UTF8.GetPreamble(), .. Encoding.UTF8.GetBytes(sb.ToString())];
    }
}
