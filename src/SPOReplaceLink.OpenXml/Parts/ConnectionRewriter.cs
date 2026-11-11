using System.Text.RegularExpressions;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.OpenXml.Text;

namespace SPOReplaceLink.OpenXml.Parts;

internal static partial class PartRewriters
{
    /// <summary>
    /// OOX-6: links in xl/connections.xml. Any element is looked at, so the x15 extensions (dataFeedPr, oledbPr…)
    /// are covered too:
    /// <list type="bullet">
    /// <item><c>connection</c> — connection string: each value that is a link (Data Source, DBQ, DefaultDir…);</item>
    /// <item><c>command</c> — SQL text: paths in `…`, […], "…", '…' and links written without quotes;</item>
    /// <item><c>sourceFile</c>, <c>odcFile</c>, <c>url</c>, <c>editPage</c> — the value is a link as a whole.</item>
    /// </list>
    /// Other attributes (names, descriptions, parameters) are left; old prefixes there go to the residual report.
    /// </summary>
    public static string RewriteConnections(string xml, LinkContext ctx)
    {
        var edits = new TextEdits();
        foreach (Match attr in Attribute().Matches(xml))
        {
            Func<string, LinkContext, IEnumerable<(int Start, int Length, string Text)>>? rewrite = LocalName(attr.Groups["name"].Value) switch
            {
                "connection" => ConnectionStringEdits,
                "command" => CommandEdits,
                "sourceFile" or "odcFile" or "url" or "editPage" => WholeValueEdits,
                _ => null,
            };
            if (rewrite is null || attr.Groups["name"].Value.StartsWith("xmlns", StringComparison.Ordinal))
                continue;
            var value = ValueGroup(attr);
            char quote = xml[value.Index - 1];
            string? replaced = RewriteEscaped(xml.Substring(value.Index, value.Length), quote, s => rewrite(s, ctx));
            if (replaced is not null)
                edits.Replace(value.Index, value.Length, replaced);
        }
        return edits.Apply(xml);
    }

    /// <summary>
    /// Applies edits made to the unescaped value back to its escaped source, so the characters around them
    /// (&amp;#13;&amp;#10; in SQL text, for one) keep their spelling. Returns null if nothing changed.
    /// </summary>
    private static string? RewriteEscaped(string escaped, char quote, Func<string, IEnumerable<(int Start, int Length, string Text)>> rewrite)
    {
        string value = XmlText.Unescape(escaped, out int[] offsets);
        var edits = new TextEdits();
        foreach (var (start, length, text) in rewrite(value))
        {
            int sourceStart = offsets[start];
            edits.Replace(sourceStart, offsets[start + length] - sourceStart, XmlText.EscapeAttribute(text, quote));
        }
        return edits.Count == 0 ? null : edits.Apply(escaped);
    }

    private static IEnumerable<(int, int, string)> WholeValueEdits(string value, LinkContext ctx)
    {
        if (RewriteLiteral(value, ctx, LinkKind.Connection, serverRelative: false) is { } replaced)
            yield return (0, value.Length, replaced);
    }

    /// <summary>Values of key=value pairs; a value may be quoted with " or ' (doubled inside) or put in {…} (ODBC).</summary>
    internal static IEnumerable<(int Start, int Length, char Quote)> ConnectionStringValues(string s)
    {
        int i = 0;
        while (i < s.Length)
        {
            int eq = s.IndexOf('=', i);
            if (eq < 0)
                yield break;
            int j = eq + 1;
            while (j < s.Length && s[j] == ' ')
                j++;
            char open = j < s.Length ? s[j] : '\0';
            char close = open switch { '"' => '"', '\'' => '\'', '{' => '}', _ => '\0' };
            if (close != '\0')
            {
                int end = ClosingIndex(s, j + 1, close);
                if (end < 0)
                    yield break;
                yield return (j + 1, end - j - 1, open);
                int semi = s.IndexOf(';', end + 1);
                i = semi < 0 ? s.Length : semi + 1;
                continue;
            }
            int stop = s.IndexOf(';', j);
            if (stop < 0)
                stop = s.Length;
            int last = stop;
            while (last > j && char.IsWhiteSpace(s[last - 1]))
                last--;
            yield return (j, last - j, '\0');
            i = stop + 1;
        }
    }

    private static IEnumerable<(int, int, string)> ConnectionStringEdits(string s, LinkContext ctx)
    {
        foreach (var (start, length, quote) in ConnectionStringValues(s))
        {
            char close = quote == '{' ? '}' : quote;
            string raw = s.Substring(start, length);
            string value = close == '\0' ? raw : raw.Replace(new string(close, 2), close.ToString());
            string? replaced = RewriteLiteral(value, ctx, LinkKind.Connection, serverRelative: false);
            if (replaced is null)
                continue;
            if (close != '\0')
                yield return (start, length, replaced.Replace(close.ToString(), new string(close, 2)));
            else if (replaced.IndexOfAny([';', '"', '\'']) >= 0 || replaced != replaced.Trim())
                yield return (start, length, "\"" + replaced.Replace("\"", "\"\"") + "\"");
            else
                yield return (start, length, replaced);
        }
    }

    /// <summary>SQL text: paths in `…`, […], "…" and '…' as a whole, links outside of them up to whitespace.</summary>
    private static IEnumerable<(int, int, string)> CommandEdits(string s, LinkContext ctx)
    {
        var result = new List<(int, int, string)>();
        void Plain(int from, int to)
        {
            if (to > from && RewriteEmbedded(s[from..to], ctx, LinkKind.Connection) is { } replaced)
                result.Add((from, to - from, replaced));
        }

        int plainStart = 0;
        int i = 0;
        while (i < s.Length)
        {
            char close = s[i] switch { '`' => '`', '[' => ']', '"' => '"', '\'' => '\'', _ => '\0' };
            if (close == '\0')
            {
                i++;
                continue;
            }
            int end = ClosingIndex(s, i + 1, close);
            if (end < 0)
                break;
            Plain(plainStart, i);
            string doubled = new(close, 2);
            string token = s.Substring(i + 1, end - i - 1).Replace(doubled, close.ToString());
            if (RewriteLiteral(token, ctx, LinkKind.Connection, serverRelative: false) is { } replaced)
                result.Add((i + 1, end - i - 1, replaced.Replace(close.ToString(), doubled)));
            i = end + 1;
            plainStart = i;
        }
        Plain(plainStart, s.Length);
        return result;
    }

    /// <summary>Index of the closing char at or after <paramref name="from"/>; a doubled one is an escape. -1 if missing.</summary>
    private static int ClosingIndex(string s, int from, char close)
    {
        int k = from;
        while (k < s.Length)
        {
            if (s[k] == close)
            {
                if (k + 1 < s.Length && s[k + 1] == close)
                {
                    k += 2;
                    continue;
                }
                return k;
            }
            k++;
        }
        return -1;
    }
}
