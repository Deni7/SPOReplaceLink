using System.Text;
using SPOReplaceLink.Core.Processing;

namespace SPOReplaceLink.Core.Reports;

/// <summary>
/// CSV reports (REP-1, REP-2, RES-2). UTF-8 with BOM so Excel shows Cyrillic correctly.
/// </summary>
public sealed class CsvReport : IDisposable
{
    private readonly StreamWriter _files;
    private readonly StreamWriter _links;
    private readonly StreamWriter _residuals;
    private readonly char _delimiter;
    private readonly object _sync = new();

    public CsvReport(string directory, char delimiter = ',')
    {
        Directory.CreateDirectory(directory);
        _delimiter = delimiter;
        _files = Open(Path.Combine(directory, "files.csv"));
        _links = Open(Path.Combine(directory, "links.csv"));
        _residuals = Open(Path.Combine(directory, "residuals.csv"));

        WriteRow(_files, "path", "format", "status", "replacements", "hyperlinks", "link_sources", "formulas",
            "cached_values", "other_external", "residuals", "changed_parts", "error_class", "error");
        WriteRow(_links, "file", "part", "kind", "action", "old_value", "new_value", "rule", "resolved_from");
        WriteRow(_residuals, "file", "part", "encoding", "match", "context");
    }

    public void Add(FileResult r)
    {
        lock (_sync)
        {
            WriteRow(_files, r.Path, r.Format, r.StatusCode, Num(r.Replacements),
                Num(r.ReplacementsOf(LinkKind.Hyperlink)), Num(r.ReplacementsOf(LinkKind.LinkSource)),
                Num(r.ReplacementsOf(LinkKind.Formula)), Num(r.ReplacementsOf(LinkKind.CachedValue)),
                Num(r.ReplacementsOf(LinkKind.OtherExternal)), Num(r.Residuals.Count),
                string.Join(' ', r.ChangedParts), r.ErrorClass?.ToString().ToLowerInvariant(), r.Error);

            foreach (var l in r.Links)
            {
                WriteRow(_links, r.Path, l.Part, Snake(l.Kind.ToString()), Snake(l.Action.ToString()),
                    l.OldValue, l.NewValue, l.RuleIndex?.ToString(), l.ResolvedFrom);
            }
            foreach (var x in r.Residuals)
                WriteRow(_residuals, r.Path, x.Part, x.Encoding, x.Match, x.Context);
        }
    }

    public void Dispose()
    {
        _files.Dispose();
        _links.Dispose();
        _residuals.Dispose();
    }

    private static StreamWriter Open(string path) => new(path, false, new UTF8Encoding(true));

    private void WriteRow(StreamWriter w, params string?[] values)
    {
        for (int i = 0; i < values.Length; i++)
        {
            if (i > 0)
                w.Write(_delimiter);
            w.Write(Escape(values[i], _delimiter));
        }
        w.Write("\r\n");
    }

    public static string Escape(string? value, char delimiter)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;
        // A leading = + - @ would be run as a formula when the report is opened in Excel.
        if (value[0] is '=' or '+' or '-' or '@')
            value = "'" + value;
        bool quote = value.IndexOfAny([delimiter, '"', '\r', '\n']) >= 0;
        return quote ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;
    }

    private static string Num(int n) => n.ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static string Snake(string pascal)
    {
        var sb = new StringBuilder(pascal.Length + 4);
        for (int i = 0; i < pascal.Length; i++)
        {
            if (i > 0 && char.IsUpper(pascal[i]))
                sb.Append('_');
            sb.Append(char.ToLowerInvariant(pascal[i]));
        }
        return sb.ToString();
    }
}
