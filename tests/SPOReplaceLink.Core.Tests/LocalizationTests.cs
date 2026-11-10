using System.Text;
using System.Text.RegularExpressions;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Core.Tests;

/// <summary>
/// The translation catalogs against the code: every translated string of the sources has an English entry with the same
/// placeholders, and no Ukrainian text is left outside L.T / L.F (comments, help section ids and lines marked
/// <c>i18n-ignore</c> excepted).
/// </summary>
public sealed partial class LocalizationTests
{
    private static string Root()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "SPOReplaceLink.sln")))
            dir = dir.Parent;
        return dir?.FullName ?? throw new InvalidOperationException("Repository root not found.");
    }

    private static IEnumerable<string> Sources() =>
        Directory.EnumerateFiles(Path.Combine(Root(), "src"), "*.*", SearchOption.AllDirectories)
            .Where(f => f.EndsWith(".cs", StringComparison.Ordinal) || f.EndsWith(".razor", StringComparison.Ordinal))
            .Where(f => !f.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.Ordinal)
                && !f.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.Ordinal)
                && !f.EndsWith($"Localization{Path.DirectorySeparatorChar}L.cs", StringComparison.Ordinal));

    [Fact]
    public void EveryTranslatedStringHasAnEnglishEntry()
    {
        var catalog = L.Catalog(L.English);
        var problems = new List<string>();
        foreach (string file in Sources())
        {
            string text = File.ReadAllText(file);
            foreach (Match m in CallRegex().Matches(text))
            {
                string key = Unescape(m.Groups["s"].Value, m.Groups["v"].Success);
                if (!CyrillicRegex().IsMatch(key))
                    continue;
                string where = $"{Path.GetFileName(file)}:{text[..m.Index].Count(c => c == '\n') + 1}";
                if (!catalog.TryGetValue(key, out var en))
                    problems.Add($"{where}: no translation for «{key}»");
                else if (m.Groups["f"].Value == "F" && !Placeholders(key).SetEquals(Placeholders(en)))
                    problems.Add($"{where}: placeholders differ in «{key}» → «{en}»");
            }
        }
        Assert.True(problems.Count == 0, string.Join("\n", problems));
    }

    [Fact]
    public void NoUkrainianTextOutsideTranslations()
    {
        var problems = new List<string>();
        foreach (string file in Sources())
        {
            string text = CallRegex().Replace(File.ReadAllText(file), "L.T(\"\"");
            string? block = null;
            int n = 0;
            foreach (string raw in text.Split('\n'))
            {
                n++;
                if (raw.Contains("i18n-ignore", StringComparison.Ordinal))
                    continue;
                string s = raw;
                if (block is not null)
                {
                    int end = s.IndexOf(block, StringComparison.Ordinal);
                    if (end < 0)
                        continue;
                    s = s[(end + 2)..];
                    block = null;
                }
                s = InlineCommentRegex().Replace(s, string.Empty);
                foreach (var (open, close) in new[] { ("/*", "*/"), ("@*", "*@") })
                {
                    int k = s.IndexOf(open, StringComparison.Ordinal);
                    if (k >= 0)
                    {
                        s = s[..k];
                        block = close;
                    }
                }
                s = LineCommentRegex().Replace(s, string.Empty);
                s = AllowedRegex().Replace(s, string.Empty);
                if (CyrillicRegex().IsMatch(s))
                    problems.Add($"{Path.GetFileName(file)}:{n}: {raw.Trim()}");
            }
        }
        Assert.True(problems.Count == 0, string.Join("\n", problems));
    }

    [Fact]
    public void LanguageCanBeChangedForAFlow()
    {
        Assert.Equal("en", L.Normalize("en-GB"));
        Assert.Null(L.Normalize("de"));
        string before = L.Language;
        using (L.Use("en"))
        {
            Assert.Equal("Help", L.T("Довідка"));
            Assert.Equal("en-US", L.Culture.Name);
        }
        Assert.Equal(before, L.Language);
        Assert.Equal("Довідка", L.Translate("Довідка", L.Ukrainian));
        Assert.Equal("Текст без перекладу", L.Translate("Текст без перекладу", L.English));
    }

    private static HashSet<string> Placeholders(string s) => PlaceholderRegex().Matches(s).Select(m => m.Groups[1].Value).ToHashSet();

    private static string Unescape(string s, bool verbatim)
    {
        if (verbatim)
            return s.Replace("\"\"", "\"", StringComparison.Ordinal);
        var sb = new StringBuilder();
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] != '\\' || i + 1 >= s.Length)
            {
                sb.Append(s[i]);
                continue;
            }
            char c = s[++i];
            switch (c)
            {
                case 'n': sb.Append('\n'); break;
                case 'r': sb.Append('\r'); break;
                case 't': sb.Append('\t'); break;
                case '0': sb.Append('\0'); break;
                case 'u':
                    sb.Append((char)Convert.ToInt32(s.Substring(i + 1, 4), 16));
                    i += 4;
                    break;
                default: sb.Append(c); break;
            }
        }
        return sb.ToString();
    }

    [GeneratedRegex(@"\bL\.(?<f>T|F|Translate)\(\s*(?<v>@)?""(?<s>(?:[^""\\]|\\.|"""")*)""")]
    private static partial Regex CallRegex();

    [GeneratedRegex(@"[Ѐ-ӿ]")]
    private static partial Regex CyrillicRegex();

    [GeneratedRegex(@"\{(\d+)(?:[,:][^}]*)?\}")]
    private static partial Regex PlaceholderRegex();

    [GeneratedRegex(@"/\*.*?\*/|@\*.*?\*@")]
    private static partial Regex InlineCommentRegex();

    [GeneratedRegex(@"(^|\s)//.*$")]
    private static partial Regex LineCommentRegex();

    [GeneratedRegex(@"Section=\{?""[^""]*""|help/[^""\s)]*")]
    private static partial Regex AllowedRegex();
}
