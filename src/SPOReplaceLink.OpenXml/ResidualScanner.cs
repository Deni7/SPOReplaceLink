using System.Text;
using System.Text.RegularExpressions;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Rules;
using SPOReplaceLink.Core.Urls;

namespace SPOReplaceLink.OpenXml;

/// <summary>
/// RES-1: finds old prefixes anywhere in a part — in any case, with / or \, raw or percent-encoded,
/// as UTF-8 text or UTF-16LE in binary parts (XLB-2).
/// </summary>
public sealed class ResidualScanner
{
    private const int ContextChars = 40;
    private readonly Regex _regex;

    public ResidualScanner(RuleSet rules)
    {
        var alternatives = new List<string>();
        foreach (var rule in rules.Rules)
        {
            alternatives.Add(Pattern(rule.MatchKey));
            // http rules: the same path written without host (/Document/…), but not inside another path.
            if (LinkClassifier.Classify(rule.MatchKey) == LinkForm.Http)
            {
                string path = PathOf(rule.MatchKey);
                if (path.Length > 1)
                    alternatives.Add(@"(?<![\p{L}\p{N}_.~%/:-])" + Pattern(path));
            }
        }
        string boundary = @"(?![\p{L}\p{N}_.~-]|%(?!2[Ff]|5[Cc]))";
        _regex = new Regex("(?:" + string.Join("|", alternatives.Distinct()) + ")" + boundary,
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled,
            TimeSpan.FromSeconds(30));
    }

    public IEnumerable<Residual> ScanText(string part, string text) => Scan(part, text, "utf-8");

    /// <summary>Binary part: looks for UTF-16LE strings (both byte alignments) and 8-bit strings.</summary>
    public IEnumerable<Residual> ScanBinary(string part, byte[] bytes)
    {
        var found = new HashSet<string>(StringComparer.Ordinal);
        var results = new List<Residual>();
        void Add(IEnumerable<Residual> items)
        {
            foreach (var r in items)
            {
                if (found.Add(r.Encoding + "|" + r.Context))
                    results.Add(r);
            }
        }
        Add(Scan(part, Encoding.Unicode.GetString(bytes), "utf-16le"));
        if (bytes.Length > 1)
            Add(Scan(part, Encoding.Unicode.GetString(bytes, 1, bytes.Length - 1), "utf-16le"));
        Add(Scan(part, Encoding.Latin1.GetString(bytes), "8-bit"));
        return results;
    }

    private IEnumerable<Residual> Scan(string part, string text, string encoding)
    {
        foreach (Match m in _regex.Matches(text))
        {
            int from = Math.Max(0, m.Index - ContextChars);
            int to = Math.Min(text.Length, m.Index + m.Length + ContextChars);
            yield return new Residual(part, encoding, Clean(m.Value), Clean(text[from..to]));
        }
    }

    /// <summary>Regex for a normalized prefix: each char matches its raw, upper-case or percent-encoded spelling.</summary>
    private static string Pattern(string normalized)
    {
        var sb = new StringBuilder();
        foreach (Rune rune in normalized.EnumerateRunes())
        {
            string ch = rune.ToString();
            if (ch == "/")
            {
                sb.Append(@"(?:/|\\|%2F|%5C)");
                continue;
            }
            string escaped = Regex.Escape(ch);
            string encoded = UrlText.EncodeUnsafe(ch);
            string upper = Rune.ToUpperInvariant(rune).ToString();
            var options = new List<string> { escaped };
            if (upper != ch)
                options.Add(Regex.Escape(upper));
            if (encoded != ch)
                options.Add(Regex.Escape(encoded));
            if (ch == " ")
                options.Add(@"\+");
            sb.Append(options.Count == 1 ? escaped : "(?:" + string.Join("|", options) + ")");
        }
        return sb.ToString();
    }

    private static string PathOf(string httpUrl)
    {
        int hostStart = httpUrl.IndexOf("://", StringComparison.Ordinal) + 3;
        int pathStart = httpUrl.IndexOf('/', hostStart);
        return pathStart < 0 ? string.Empty : httpUrl[pathStart..];
    }

    private static string Clean(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (char c in s)
            sb.Append(char.IsControl(c) || char.IsSurrogate(c) ? '·' : c);
        return sb.ToString();
    }
}
