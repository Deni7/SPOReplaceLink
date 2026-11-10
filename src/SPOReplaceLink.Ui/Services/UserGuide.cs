using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Markdig;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Ui.Services;

/// <summary>
/// docs/USER_GUIDE.md, embedded in the app and split into sections for the help pages: parts (#) group the
/// sections (##), a section with subsections (###) gets one page for its introduction and one per subsection.
/// Section ids are the GitHub anchors of the headings, so links inside the guide lead to the right page.
/// The guide in another language (docs/USER_GUIDE.{lang}.md) has the same headings in the same order: its sections take
/// the ids of the Ukrainian ones, so the "?" buttons and links of the app lead to the same section in every language.
/// </summary>
public sealed partial class UserGuide
{
    public sealed record Section(string Id, string Title, int Level, string Group, string? ParentId, string Html, string Text)
    {
        public List<Section> Children { get; } = [];
    }

    public sealed record Part(string Title, IReadOnlyList<Section> Sections);

    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, UserGuide> Embedded = new(StringComparer.Ordinal);

    /// <summary>The guide in the current language (the Ukrainian one when there is no translation).</summary>
    public static UserGuide Instance => For(L.Language);

    public static UserGuide For(string language) => Embedded.GetOrAdd(language, lang =>
    {
        var source = new UserGuide(Resource("USER_GUIDE.md")!, L.Ukrainian);
        if (lang == L.Ukrainian || Resource($"USER_GUIDE.{lang}.md") is not { } translated)
            return source;
        return new UserGuide(translated, lang, source.HeadingIds);
    });

    private static string? Resource(string name)
    {
        using var stream = typeof(UserGuide).Assembly.GetManifestResourceStream(name);
        if (stream is null)
            return name == "USER_GUIDE.md" ? throw new InvalidOperationException(L.T("Інструкцію не вбудовано в застосунок.")) : null;
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    /// <summary>Headings (#, ##, ###) after the title of the document, in order, outside code blocks.</summary>
    public static IReadOnlyList<(int Level, string Title)> Headings(string markdown)
    {
        var result = new List<(int, string)>();
        bool inCode = false;
        bool first = true;
        foreach (string raw in markdown.Replace("\r\n", "\n").Split('\n'))
        {
            if (raw.TrimStart().StartsWith("```", StringComparison.Ordinal))
                inCode = !inCode;
            if (!inCode && HeadingRegex().Match(raw.TrimEnd()) is { Success: true } m)
            {
                if (first && m.Groups[1].Length == 1 && result.Count == 0)
                {
                    first = false;
                    continue;
                }
                first = false;
                result.Add((m.Groups[1].Length, m.Groups[2].Value.Trim()));
            }
        }
        return result;
    }

    /// <summary>Ids of the headings in order (see <see cref="Headings"/>).</summary>
    public IReadOnlyList<string> HeadingIds { get; }

    /// <summary>Language of this guide.</summary>
    public string Language { get; }

    private readonly Dictionary<string, string> _anchors = new(StringComparer.Ordinal);

    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder().UsePipeTables().UseEmphasisExtras().Build();

    private readonly Dictionary<string, Section> _byId = new(StringComparer.Ordinal);

    /// <param name="ids">Ids for the headings in order (those of the source guide); none — the anchors of the headings.</param>
    public UserGuide(string markdown, string language = L.Ukrainian, IReadOnlyList<string>? ids = null)
    {
        Language = language;
        var headings = Headings(markdown);
        HeadingIds = headings.Select((h, i) => ids is not null && i < ids.Count ? ids[i] : Slug(h.Title)).ToList();
        for (int i = 0; i < headings.Count; i++)
            _anchors.TryAdd(Slug(headings[i].Title), HeadingIds[i]);
        int index = 0;
        string? currentId = null;
        var parts = new List<(string Title, List<Section> Sections)>();
        var body = new StringBuilder();
        string? docTitle = null;
        (string Title, int Level)? current = null; // null — the text before the first section (introduction)
        bool skip = false; // the table of contents: the app has its own navigation
        bool inCode = false;
        Section? parent = null;

        void Flush()
        {
            string text = body.ToString().Trim().Trim('-').Trim();
            body.Clear();
            if (skip)
                return;
            if (current is not { } h)
            {
                if (text.Length > 0)
                    Intro = new Section("intro", L.Translate("Про інструкцію", language), 2, string.Empty, null, ToHtml(text), Plain(text));
                return;
            }
            string group = h.Level == 1 ? h.Title : parts.Count > 0 ? parts[^1].Title : string.Empty;
            var section = new Section(currentId!, h.Title, h.Level, group, h.Level == 3 ? parent?.Id : null, ToHtml(text), Plain(h.Title + "\n" + text));
            if (h.Level == 1)
            {
                parts.Add((h.Title, []));
                // A part's own text (before its first section) becomes its first page.
                if (text.Length > 0)
                    parts[^1].Sections.Add(section);
            }
            else if (h.Level == 2)
            {
                if (parts.Count == 0)
                    parts.Add((string.Empty, []));
                parts[^1].Sections.Add(section);
                parent = section;
            }
            else
            {
                parent?.Children.Add(section);
            }
            _byId.TryAdd(section.Id, section);
        }

        foreach (string raw in markdown.Replace("\r\n", "\n").Split('\n'))
        {
            if (raw.TrimStart().StartsWith("```", StringComparison.Ordinal))
                inCode = !inCode;
            if (!inCode && HeadingRegex().Match(raw.TrimEnd()) is { Success: true } m)
            {
                int level = m.Groups[1].Length;
                string text = m.Groups[2].Value.Trim();
                if (level == 1 && docTitle is null && current is null && body.ToString().Trim().Length == 0)
                {
                    docTitle = text;
                    continue;
                }
                Flush();
                current = (text, level);
                currentId = index < HeadingIds.Count ? HeadingIds[index++] : Slug(text);
                skip = level == 2 && text is "Зміст" or "Contents" or "Table of contents"; // i18n-ignore: heading of the Ukrainian guide
                continue;
            }
            body.AppendLine(raw);
        }
        Flush();

        Title = docTitle ?? L.Translate("Інструкція", language);
        if (Intro is not null)
            _byId[Intro.Id] = Intro;
        Parts = parts.Select(p => new Part(p.Title, p.Sections)).ToList();
        Pages = (Intro is null ? [] : new[] { Intro })
            .Concat(Parts.SelectMany(p => p.Sections.SelectMany(s => s.Children.Count == 0 || HasOwnText(s) ? new[] { s }.Concat(s.Children) : s.Children)))
            .ToList();
    }

    public string Title { get; }
    public Section? Intro { get; private set; }

    public IReadOnlyList<Part> Parts { get; }

    /// <summary>Pages in reading order (for "previous / next").</summary>
    public IReadOnlyList<Section> Pages { get; }

    public Section? Find(string? id) => id is not null && _byId.TryGetValue(id, out var s) ? s : null;

    public IEnumerable<string> Ids => _byId.Keys;

    /// <summary>Sections whose title or text contains all the words of the query, titles first.</summary>
    public IReadOnlyList<Section> Search(string query)
    {
        var words = query.ToLower(CultureInfo.CurrentCulture).Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (words.Length == 0)
            return [];
        return Pages
            .Where(s => words.All(w => s.Text.Contains(w, StringComparison.Ordinal)))
            .OrderByDescending(s => words.Count(w => s.Title.Contains(w, StringComparison.OrdinalIgnoreCase)))
            .ToList();
    }

    private static bool HasOwnText(Section s) => s.Html.Length > 0;

    /// <summary>GitHub heading anchor: lower case, punctuation removed, spaces to hyphens.</summary>
    public static string Slug(string heading)
    {
        var sb = new StringBuilder();
        foreach (char c in heading.Trim().ToLowerInvariant())
        {
            if (char.IsLetterOrDigit(c) || c is '-' or '_')
                sb.Append(c);
            else if (c == ' ')
                sb.Append('-');
        }
        return sb.ToString();
    }

    private string ToHtml(string markdown)
    {
        string html = Markdown.ToHtml(markdown.Trim().TrimEnd('-').Trim(), Pipeline);
        // Links inside the guide open the help page of the section; links to other documents of the repository are
        // not in the app, so only their text is kept.
        html = AnchorRegex().Replace(html, m => $"href=\"help/{(_anchors.TryGetValue(m.Groups[1].Value, out var id) ? id : m.Groups[1].Value)}\"");
        html = RepoLinkRegex().Replace(html, "$1");
        html = ExternalLinkRegex().Replace(html, m => $"{m.Value} target=\"_blank\"");
        return html;
    }

    private static string Plain(string markdown) =>
        Regex.Replace(markdown, @"[`*|#>\[\]]|\(#[^)]*\)", " ").ToLower(CultureInfo.CurrentCulture);

    [GeneratedRegex(@"^(#{1,3})\s+(.+)$")]
    private static partial Regex HeadingRegex();

    [GeneratedRegex("href=\"#([^\"]+)\"")]
    private static partial Regex AnchorRegex();

    [GeneratedRegex("<a href=\"\\.\\./[^\"]*\"[^>]*>(.*?)</a>")]
    private static partial Regex RepoLinkRegex();

    [GeneratedRegex("href=\"https?://[^\"]+\"")]
    private static partial Regex ExternalLinkRegex();
}
