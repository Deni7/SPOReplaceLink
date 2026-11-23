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
    /// <summary>Window of a binary part scanned at once; a part above it is read piece by piece.</summary>
    internal const int DefaultChunkBytes = 32 * 1024 * 1024;
    private readonly Regex _regex;
    /// <summary>Longest text a match can take: each char of a key spelled as up to four percent-encoded bytes.</summary>
    private readonly int _maxMatch;

    public ResidualScanner(RuleSet rules)
    {
        // Alternatives share long prefixes (host, /sites/…): they are merged into a trie so that each text
        // position walks one branch instead of trying every rule — with ~100 rules a flat alternation
        // takes tens of seconds on a large sheet.
        var trie = new Node();
        int longest = 0;
        foreach (var rule in rules.Rules)
        {
            longest = Math.Max(longest, rule.MatchKey.Length);
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
        _maxMatch = 12 * longest + 16;
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
    public IEnumerable<Residual> ScanBinary(string part, byte[] bytes) => ScanBinary(part, new MemoryStream(bytes, writable: false));

    /// <summary>
    /// The same for a part read from a stream: windows of <paramref name="chunkBytes"/> overlapped on both sides by more
    /// than a match and its context, so a part of any size is scanned without holding it whole; a match counts in the
    /// window where it starts.
    /// </summary>
    public IReadOnlyList<Residual> ScanBinary(string part, Stream stream, int chunkBytes = DefaultChunkBytes)
    {
        var found = new HashSet<string>(StringComparer.Ordinal);
        var results = new List<Residual>();
        // Even sizes keep UTF-16 alignment of every window the same as in the whole part.
        int overlap = 2 * (_maxMatch + 2 * ContextChars + 8);
        chunkBytes = Math.Max(2, chunkBytes & ~1);
        var window = new byte[overlap + chunkBytes + overlap];
        long windowStart = 0;
        int filled = 0;
        bool eof = false;
        for (long own = 0; ; own += chunkBytes)
        {
            long keepFrom = Math.Max(0, own - overlap);
            int drop = (int)(keepFrom - windowStart);
            if (drop > 0)
            {
                Buffer.BlockCopy(window, drop, window, 0, filled - drop);
                filled -= drop;
                windowStart = keepFrom;
            }
            int want = (int)(own + chunkBytes + overlap - windowStart);
            while (filled < want && !eof)
            {
                int n = stream.Read(window, filled, want - filled);
                if (n == 0)
                    eof = true;
                else
                    filled += n;
            }

            long ownEnd = own + chunkBytes;
            void Add(string text, string encoding, int origin, int step)
            {
                foreach (var r in Scan(part, text, encoding, i => windowStart + origin + (long)step * i is var at && at >= own && at < ownEnd))
                {
                    if (found.Add(r.Encoding + "|" + r.Context))
                        results.Add(r);
                }
            }
            Add(Encoding.Unicode.GetString(window, 0, filled), "utf-16le", 0, 2);
            if (filled > 1)
                Add(Encoding.Unicode.GetString(window, 1, filled - 1), "utf-16le", 1, 2);
            Add(Encoding.Latin1.GetString(window, 0, filled), "8-bit", 0, 1);

            if (eof && ownEnd >= windowStart + filled)
                return results;
        }
    }

    private IEnumerable<Residual> Scan(string part, string text, string encoding, Func<int, bool>? counts = null)
    {
        foreach (Match m in _regex.Matches(text))
        {
            if (counts is not null && !counts(m.Index))
                continue;
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
