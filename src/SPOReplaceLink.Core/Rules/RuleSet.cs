using SPOReplaceLink.Core.Urls;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Core.Rules;

public sealed record RuleSetBuild(RuleSet? RuleSet, IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsValid => RuleSet is not null;
}

/// <summary>
/// Validated set of rules. A link is rewritten in one pass by the rule with the longest matching
/// old prefix (RUL-2); the result is never matched again. Rule order has no effect.
/// </summary>
public sealed class RuleSet
{
    private readonly CompiledRule[] _byLength;

    private RuleSet(CompiledRule[] rules, string? baseHost)
    {
        _byLength = rules
            .OrderByDescending(r => r.MatchKey.Length)
            .ThenBy(r => r.MatchKey, StringComparer.Ordinal)
            .ToArray();
        Rules = rules.OrderBy(r => r.Index).ToArray();
        BaseHost = baseHost;
    }

    public IReadOnlyList<CompiledRule> Rules { get; }

    /// <summary>https://tenant.sharepoint.com — used for server-relative rules and links.</summary>
    public string? BaseHost { get; }

    public static RuleSetBuild Build(IReadOnlyList<RewriteRule> rules, string? baseHost)
    {
        var issues = new List<ValidationIssue>();
        baseHost = string.IsNullOrWhiteSpace(baseHost) ? null : baseHost.Trim().TrimEnd('/');
        if (baseHost is not null && !IsValidBaseHost(baseHost))
        {
            issues.Add(new(IssueSeverity.Error, L.F("baseHost «{0}» має бути адресою виду https://tenant.sharepoint.com без шляху.", baseHost)));
            baseHost = null;
        }
        if (rules.Count == 0)
            issues.Add(new(IssueSeverity.Error, L.T("Набір правил порожній.")));

        var compiled = new List<CompiledRule>();
        for (int i = 0; i < rules.Count; i++)
        {
            var rule = rules[i];
            string? oldAbs = ToAbsolute(rule.Old, baseHost, i + 1, L.T("старий"), issues);
            string? newAbs = ToAbsolute(rule.New, baseHost, i + 1, L.T("новий"), issues);
            if (oldAbs is not null && newAbs is not null)
                compiled.Add(new CompiledRule(i + 1, rule, oldAbs, newAbs));
        }

        var effective = RemoveDuplicatesAndIdentities(compiled, issues);
        if (issues.Any(x => x.Severity == IssueSeverity.Error))
            return new RuleSetBuild(null, issues);

        var set = new RuleSet(effective.ToArray(), baseHost);
        set.CheckIdempotency(issues);
        if (issues.Any(x => x.Severity == IssueSeverity.Error))
            return new RuleSetBuild(null, issues);
        return new RuleSetBuild(set, issues);
    }

    /// <param name="documentFolder">Where the document lived before migration; needed for relative links.</param>
    public RewriteResult Rewrite(string link, string? documentFolder = null)
    {
        var form = LinkClassifier.Classify(link);
        string effective;
        switch (form)
        {
            case LinkForm.Empty or LinkForm.Internal:
                return new RewriteResult(RewriteOutcome.NotApplicable, link);
            case LinkForm.ServerRelative:
                // Relative to the root of the document's location (host, server or drive — Excel writes
                // "/Users/…" for a file on the same drive); without it, to the tenant host.
                string? root = (documentFolder is null ? null : RelativeLinkResolver.RootOf(documentFolder)) ?? BaseHost;
                if (root is null)
                    return new RewriteResult(RewriteOutcome.Unresolved, link);
                effective = root + link;
                break;
            case LinkForm.Relative:
                string? resolved = documentFolder is null ? null : RelativeLinkResolver.Resolve(documentFolder, link);
                if (resolved is null)
                    return new RewriteResult(RewriteOutcome.Unresolved, link);
                effective = resolved;
                break;
            default:
                effective = link;
                break;
        }

        var normalized = NormalizedLink.Create(effective);
        var rule = FindLongest(normalized.Value);
        string? resolvedFrom = ReferenceEquals(effective, link) ? null : effective;
        if (rule is null)
            return new RewriteResult(RewriteOutcome.NoRule, link, ResolvedFrom: resolvedFrom);

        int split = normalized.OriginalIndex(rule.MatchKey.Length);
        string output = Render(rule, effective, LinkClassifier.Classify(effective), split);
        if (output == link)
            return new RewriteResult(RewriteOutcome.AlreadyCurrent, link, link, rule, resolvedFrom);
        return new RewriteResult(RewriteOutcome.Replaced, link, output, rule, resolvedFrom);
    }

    /// <summary>The rule with the longest old prefix that matches at a segment boundary (RUL-4).</summary>
    public CompiledRule? FindLongest(string normalized)
    {
        foreach (var rule in _byLength)
        {
            if (Matches(rule.MatchKey, normalized))
                return rule;
        }
        return null;
    }

    /// <summary>
    /// Where a location in the new place used to be: the rule whose new prefix matches is applied backwards.
    /// Null if no rule matches or several rules lead to the same new prefix from different old ones.
    /// </summary>
    public string? Reverse(string currentLocation)
    {
        var normalized = NormalizedLink.Create(currentLocation);
        var candidates = Rules.Where(r => Matches(r.NewKey, normalized.Value)).ToList();
        if (candidates.Count == 0)
            return null;
        int longest = candidates.Max(r => r.NewKey.Length);
        var best = candidates.Where(r => r.NewKey.Length == longest).ToList();
        if (best.Select(r => r.MatchKey).Distinct(StringComparer.Ordinal).Count() != 1)
            return null;
        string tail = currentLocation[normalized.OriginalIndex(longest)..];
        return CompiledRule.TrimTrailingSlash(best[0].EffectiveOld) + tail;
    }

    /// <summary>All rules whose old prefix matches, longest first — for <c>rules test</c>.</summary>
    public IEnumerable<CompiledRule> FindAll(string normalized) =>
        _byLength.Where(r => Matches(r.MatchKey, normalized));

    internal static bool Matches(string key, string normalized) =>
        normalized.StartsWith(key, StringComparison.Ordinal) && IsBoundary(normalized, key.Length);

    private static bool IsBoundary(string s, int index) =>
        index == s.Length || s[index] is '/' or '?' or '#' || (index > 0 && s[index - 1] == '/');

    private static string Render(CompiledRule rule, string effective, LinkForm effectiveForm, int split)
    {
        string tail = effective[split..];
        // A root rule (C:\ or \\server\) leaves the tail without its leading separator.
        if (tail.Length > 0 && tail[0] is not ('/' or '\\' or '?' or '#'))
            tail = "/" + tail;
        if (tail.Length > 0 && LinkClassifier.IsSlash(tail[0]) && LinkClassifier.IsSlash(rule.NewCanonical[^1]))
            tail = tail[1..];
        if (rule.NewForm == LinkForm.Http)
        {
            // RUL-5: the tail is copied verbatim and the new prefix follows the link's own spelling.
            // A file system path becomes a URL in the style Excel writes (spaces %20, Cyrillic as is).
            if (effectiveForm != LinkForm.Http)
            {
                var (_, nonAscii) = UrlText.DetectStyle(effective);
                tail = UrlText.Encode(tail.Replace('\\', '/'), encodeAscii: true, nonAscii);
                return UrlText.Encode(rule.NewCanonical, encodeAscii: true, nonAscii) + tail;
            }
            var (ascii, encodeNonAscii) = UrlText.DetectStyle(effective);
            return UrlText.Encode(rule.NewCanonical, ascii, encodeNonAscii) + tail;
        }

        // New location is a file system path: write it decoded with its own separator.
        if (effectiveForm == LinkForm.Http)
            tail = UrlText.Decode(tail);
        int lastSlash = rule.NewCanonical.LastIndexOfAny(['/', '\\']);
        char separator = lastSlash >= 0 ? rule.NewCanonical[lastSlash] : '\\';
        tail = tail.Replace(separator == '\\' ? '/' : '\\', separator);
        // Excel writes some paths as file:///C:\… — keep that spelling for the new path too.
        string scheme = effectiveForm == LinkForm.File && rule.NewForm is LinkForm.Drive or LinkForm.Unc ? "file:///" : string.Empty;
        return scheme + rule.NewCanonical + tail;
    }

    private static string? ToAbsolute(string value, string? baseHost, int index, string which, List<ValidationIssue> issues)
    {
        value = value?.Trim() ?? string.Empty;
        var form = LinkClassifier.Classify(value);
        if (form == LinkForm.Empty)
        {
            issues.Add(new(IssueSeverity.Error, L.F("{0} префікс порожній.", Capitalize(which)), index));
            return null;
        }
        if (LinkClassifier.IsAbsolute(form))
            return value;
        if (form == LinkForm.ServerRelative)
        {
            if (baseHost is not null)
                return baseHost + value.Replace('\\', '/');
            issues.Add(new(IssueSeverity.Error,
                L.F("{0} префікс «{1}» без хоста: задайте baseHost у профілі або вкажіть повну адресу.", Capitalize(which), value), index));
            return null;
        }
        issues.Add(new(IssueSeverity.Error,
            L.F("{0} префікс «{1}» має бути повною адресою (https://…, \\\\сервер\\…, C:\\…) або шляхом від кореня сайту (/sites/…).", Capitalize(which), value), index));
        return null;
    }

    private static List<CompiledRule> RemoveDuplicatesAndIdentities(List<CompiledRule> compiled, List<ValidationIssue> issues)
    {
        var result = new List<CompiledRule>();
        foreach (var group in compiled.GroupBy(r => r.MatchKey, StringComparer.Ordinal))
        {
            var first = group.First();
            foreach (var other in group.Skip(1))
            {
                if (other.NewKey == first.NewKey)
                    issues.Add(new(IssueSeverity.Warning, L.F("Дублює правило {0}; ігнорується.", first.Index), other.Index));
                else
                    issues.Add(new(IssueSeverity.Error,
                        L.F("Конфлікт з правилом {0}: однаковий старий префікс «{1}», різні нові («{2}» і «{3}»).", first.Index, first.Source.Old, first.Source.New, other.Source.New),
                        other.Index));
            }
            if (first.MatchKey == first.NewKey)
            {
                issues.Add(new(IssueSeverity.Warning, L.T("Старий і новий префікси однакові; правило ігнорується."), first.Index));
                continue;
            }
            result.Add(first);
        }
        return result;
    }

    /// <summary>
    /// RUL-8: a link rewritten once must not match any rule again, otherwise a second run would
    /// change it again (this also catches cycles A→B, B→A).
    /// </summary>
    private void CheckIdempotency(List<ValidationIssue> issues)
    {
        foreach (var r in Rules)
        {
            var again = FindLongest(r.NewKey);
            if (again is not null)
            {
                string what = again == r ? L.T("цим самим правилом") : L.F("правилом {0} («{1}»)", again.Index, again.Source.Old);
                issues.Add(new(IssueSeverity.Error,
                    L.F("Новий префікс «{0}» знову збігається з {1}: повторний прогін змінить посилання ще раз.", r.Source.New, what), r.Index));
                continue;
            }

            // Outputs "New + /X…" match a rule whose old prefix is "New/X". That is fine only when a more
            // specific rule already takes the inputs "Old/X…", so this rule never produces such output.
            foreach (var s in Rules)
            {
                if (s == r || s.MatchKey.Length <= r.NewKey.Length || !Matches(r.NewKey, s.MatchKey))
                    continue;
                string extra = s.MatchKey[r.NewKey.Length..];
                if (FindLongest(r.MatchKey + extra) == r)
                {
                    issues.Add(new(IssueSeverity.Error,
                        L.F("Посилання «{0}{1}…» стануть «{2}{1}…», а це старий префікс правила {3}: повторний прогін змінить їх ще раз. Додайте правило для «{0}{1}» або змініть правило {3}.",
                            r.Source.Old.TrimEnd('/', '\\'), extra, r.Source.New.TrimEnd('/', '\\'), s.Index),
                        r.Index));
                }
            }
        }
    }

    private static bool IsValidBaseHost(string baseHost) =>
        Uri.TryCreate(baseHost, UriKind.Absolute, out var uri)
        && uri.Scheme is "https" or "http"
        && uri.AbsolutePath == "/"
        && string.IsNullOrEmpty(uri.Query);

    private static string Capitalize(string s) => char.ToUpperInvariant(s[0]) + s[1..];
}
