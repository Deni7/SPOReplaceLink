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
        // Alternatives share long prefixes (host, /sites/…): they are merged into a trie so that each text
        // position walks one branch instead of trying every rule — with ~100 rules a flat alternation
        // takes tens of seconds on a large sheet.
        var trie = new Node();
        foreach (var rule in rules.Rules)
        {
            trie.Add(Tokens(rule.MatchKey));
            // http rules: the same path written without host (/Document/…), but not inside another path.
            if (LinkClassifier.Classify(rule.MatchKey) == LinkForm.Http)
            {
                string path = PathOf(rule.MatchKey);
                if (path.Length > 1)
                    trie.Add([@"(?<![\p{L}\p{N}_.~%/:-])", .. Tokens(path)]);
            }
        }
        string boundary = @"(?![\p{L}\p{N}_.~-]|%(?!2[Ff]|5[Cc]))";
        // No nested quantifiers: matching is linear in the text, so a timeout would only fail big files.
        // Not Compiled: for a pattern this size the generated method is poorly optimized by the JIT and runs
        // 40-70x slower than the interpreter (10-20 s vs 0.3 s on a 30 MB sheet).
        _regex = new Regex("(?:" + trie.ToPattern() + ")" + boundary,
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
            Regex.InfiniteMatchTimeout);
    }

    /// <summary>Trie of regex tokens; longer continuations come before the end so the longest prefix wins.</summary>
    private sealed class Node
    {
        private readonly List<(string Token, Node Next)> _children = [];
        private bool _end;

        public void Add(IEnumerable<string> tokens)
        {
            var node = this;
            foreach (string token in tokens)
            {
                var next = node._children.FirstOrDefault(c => c.Token == token).Next;
                if (next is null)
                {
                    next = new Node();
                    node._children.Add((token, next));
                }
                node = next;
            }
            node._end = true;
        }

        public string ToPattern()
        {
            var sb = new StringBuilder();
            Append(sb);
            return sb.ToString();
        }

        private void Append(StringBuilder sb)
        {
            var node = this;
            // A chain of single children is written as a plain sequence (iteratively: keys can be long).
            while (node._children.Count == 1 && !node._end)
            {
                sb.Append(node._children[0].Token);
                node = node._children[0].Next;
            }
            if (node._children.Count == 0)
                return;
            sb.Append("(?:");
            for (int i = 0; i < node._children.Count; i++)
            {
                if (i > 0)
                    sb.Append('|');
                sb.Append(node._children[i].Token);
                node._children[i].Next.Append(sb);
            }
            if (node._end)
                sb.Append('|');
            sb.Append(')');
        }
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

    /// <summary>Regex tokens for a normalized prefix: each char matches its raw, upper-case or percent-encoded spelling.</summary>
    private static IEnumerable<string> Tokens(string normalized)
    {
        foreach (Rune rune in normalized.EnumerateRunes())
        {
            string ch = rune.ToString();
            if (ch == "/")
            {
                yield return @"(?:/|\\|%2F|%5C)";
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
            yield return options.Count == 1 ? escaped : "(?:" + string.Join("|", options) + ")";
        }
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
