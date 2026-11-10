using System.Text.Json;
using System.Text.Json.Serialization;
using SPOReplaceLink.Core.Rules;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Core.Config;

/// <summary>
/// Run profile (CFG-1): targets, filters, replacement places, performance, rule templates and connections.
/// The same JSON is stored by the host, exported from the GUI and accepted by the CLI.
/// </summary>
public sealed class Profile
{
    public string? Name { get; set; }

    /// <summary>https://tenant.sharepoint.com — for server-relative rules and links (/sites/…).</summary>
    public string? BaseHost { get; set; }

    /// <summary>Names of rule templates stored by the host; their rules come before <see cref="Rules"/> (GUI-24).</summary>
    public List<string> Templates { get; set; } = [];

    public List<RuleEntry> Rules { get; set; } = [];

    public PlacesOptions Places { get; set; } = new();

    public FileFilters Filters { get; set; } = new();

    /// <summary>CFG-2: site + library + optional folder.</summary>
    public List<Target> Targets { get; set; } = [];

    /// <summary>Names of connections (app registrations) stored by the host. More than one needs <see cref="Pool"/>.</summary>
    public List<string> Connections { get; set; } = [];

    public PoolOptions Pool { get; set; } = new();

    public PerformanceOptions Performance { get; set; } = new();

    public RetryOptions Retry { get; set; } = new();

    /// <summary>Copy with template rules put in place of template names — the form a run works with (CFG-8).</summary>
    public Profile ResolveTemplates(IReadOnlyDictionary<string, Template> templates)
    {
        var copy = Parse(ToJson());
        var rules = new List<RuleEntry>();
        foreach (string name in Templates)
        {
            if (!templates.TryGetValue(name, out var template))
                throw new ProfileException(L.F("Шаблон «{0}» не знайдено.", name));
            rules.AddRange(template.Rules.Select(r => new RuleEntry { Old = r.Old, New = r.New, Name = r.Name ?? name }));
        }
        rules.AddRange(copy.Rules);
        copy.Rules = rules;
        copy.Templates = [];
        return copy;
    }

    /// <summary>CFG-7 for runs in SharePoint: everything <see cref="Validate"/> checks plus targets, connections, limits.</summary>
    public ProfileValidation ValidateForRun(IReadOnlyCollection<Connection> connections)
    {
        var basic = Validate();
        var issues = basic.Issues.ToList();
        if (Targets.Count == 0)
            issues.Add(new(IssueSeverity.Error, L.T("targets: не задано жодної цілі.")));
        for (int i = 0; i < Targets.Count; i++)
            issues.AddRange(Targets[i].Validate(i + 1));
        if (Connections.Count == 0)
            issues.Add(new(IssueSeverity.Error, L.T("connections: не задано жодного підключення.")));
        foreach (string name in Connections.Where(n => connections.All(c => !string.Equals(c.Name, n, StringComparison.OrdinalIgnoreCase))))
            issues.Add(new(IssueSeverity.Error, L.F("connections: підключення «{0}» не знайдено.", name)));
        foreach (var c in connections)
            issues.AddRange(c.Validate());
        if (Connections.Count > 1 && !Pool.Enabled)
            issues.Add(new(IssueSeverity.Error,
                L.T("Кілька підключень працюють лише як пул: увімкніть pool.enabled. Microsoft не рекомендує кілька застосунків для однієї задачі — пул ділить тенантний бюджет з користувачами.")));
        if (Pool.Enabled && Connections.Count > 1)
            issues.Add(new(IssueSeverity.Warning,
                L.F("Увімкнено пул з {0} застосунків. Сумарне навантаження обмежене {1}% тенантного бюджету.", Connections.Count, Pool.TenantCeilingPercent)));
        issues.AddRange(Pool.Validate());
        issues.AddRange(Performance.Validate());
        issues.AddRange(Retry.Validate());
        bool ok = basic.RuleSet is not null && issues.All(x => x.Severity != IssueSeverity.Error);
        return new ProfileValidation(ok ? basic.RuleSet : null, issues);
    }

    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        WriteIndented = true,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
    };

    public static Profile Parse(string json)
    {
        try
        {
            return JsonSerializer.Deserialize<Profile>(json, JsonOptions)
                ?? throw new ProfileException(L.T("Профіль порожній."));
        }
        catch (JsonException ex)
        {
            string where = ex.LineNumber is long line ? L.F(" (рядок {0})", line + 1) : string.Empty;
            throw new ProfileException(L.F("Профіль не розібрано{0}: {1}", where, ex.Message), ex);
        }
    }

    public static Profile Load(string path)
    {
        if (!File.Exists(path))
            throw new ProfileException(L.F("Файл профілю не знайдено: {0}", path));
        return Parse(File.ReadAllText(path));
    }

    public string ToJson() => JsonSerializer.Serialize(this, JsonOptions);

    public IReadOnlyList<RewriteRule> ToRules() =>
        Rules.Select(r => new RewriteRule(r.Old ?? string.Empty, r.New ?? string.Empty, r.Name)).ToList();

    /// <summary>CFG-7: checks the whole profile, including the rule set (RUL-8).</summary>
    public ProfileValidation Validate()
    {
        var issues = new List<ValidationIssue>();
        if (Templates.Count > 0)
            issues.Add(new(IssueSeverity.Error, L.T("Профіль посилається на шаблони (templates); їх підставляє хост — для локальної роботи перенесіть правила в rules.")));
        issues.AddRange(Places.Validate());
        issues.AddRange(Filters.Validate());
        var build = RuleSet.Build(ToRules(), BaseHost);
        issues.AddRange(build.Issues);
        return new ProfileValidation(build.RuleSet, issues);
    }
}

public sealed record ProfileValidation(RuleSet? RuleSet, IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsValid => RuleSet is not null && Issues.All(i => i.Severity != IssueSeverity.Error);
}

public sealed class RuleEntry
{
    public string? Old { get; set; }

    public string? New { get; set; }

    public string? Name { get; set; }
}

/// <summary>Where links are replaced (CFG-4).</summary>
public sealed class PlacesOptions
{
    /// <summary>OOX-1: hyperlinks on cells and shapes.</summary>
    public bool Hyperlinks { get; set; } = true;

    /// <summary>OOX-2, OOX-3: external workbooks (Edit Links).</summary>
    public bool LinkSources { get; set; } = true;

    /// <summary>OOX-4: URLs in formula text (HYPERLINK, INDIRECT) and their cached results.</summary>
    public bool Formulas { get; set; } = true;

    /// <summary>Other external relationships: linked OLE objects, linked pictures.</summary>
    public bool OtherExternal { get; set; } = true;

    /// <summary>OOX-5 (P2).</summary>
    public bool PivotCaches { get; set; }

    /// <summary>OOX-6 (P2).</summary>
    public bool Connections { get; set; }

    /// <summary>OOX-8 (P3).</summary>
    public bool PowerQuery { get; set; }

    /// <summary>OOX-9 (P3).</summary>
    public bool CellText { get; set; }

    internal IEnumerable<ValidationIssue> Validate()
    {
        var notYet = new (bool On, string Name)[]
        {
            (PivotCaches, "pivotCaches"), (Connections, "connections"),
            (PowerQuery, "powerQuery"), (CellText, "cellText"),
        };
        foreach (var (on, name) in notYet.Where(x => x.On))
        {
            yield return new(IssueSeverity.Error,
                L.F("places.{0}: заміна в цьому місці ще не реалізована. Вимкніть її; залишки там потраплять у звіт.", name));
        }
        if (!Hyperlinks && !LinkSources && !Formulas && !OtherExternal)
            yield return new(IssueSeverity.Error, L.T("places: вимкнено всі місця заміни."));
    }
}

public sealed class FileFilters
{
    public static readonly string[] SupportedExtensions = [".xlsx", ".xlsm", ".xlsb", ".xls"];

    public List<string> Extensions { get; set; } = [.. SupportedExtensions];

    /// <summary>Larger files get status skipped:too_large. 0 — no limit.</summary>
    public int MaxFileSizeMb { get; set; } = 500;

    /// <summary>Smaller files are left out (skipped:filtered). 0 — no limit.</summary>
    public int MinFileSizeKb { get; set; }

    /// <summary>Path masks relative to the target (* — within a folder, ** — any depth). Empty — everything.</summary>
    public List<string> Include { get; set; } = [];

    public List<string> Exclude { get; set; } = [];

    /// <summary>Only files changed on or after this date.</summary>
    public DateTimeOffset? ModifiedAfter { get; set; }

    /// <summary>Only files changed before this date.</summary>
    public DateTimeOffset? ModifiedBefore { get; set; }

    public bool Accepts(string path) =>
        Extensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase)
        && !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal);

    /// <summary>CFG-3 for a file in SharePoint; <paramref name="relativePath"/> uses '/' and is relative to the target.</summary>
    public FilterVerdict Check(string relativePath, long size, DateTimeOffset? modified)
    {
        if (!Accepts(relativePath))
            return FilterVerdict.Extension;
        if (Include.Count > 0 && !Include.Any(m => PathMask.IsMatch(m, relativePath)))
            return FilterVerdict.Path;
        if (Exclude.Any(m => PathMask.IsMatch(m, relativePath)))
            return FilterVerdict.Path;
        if (modified is { } m2 && ((ModifiedAfter is { } a && m2 < a) || (ModifiedBefore is { } b && m2 >= b)))
            return FilterVerdict.Date;
        if (MaxFileSizeMb > 0 && size > MaxFileSizeMb * 1024L * 1024L)
            return FilterVerdict.TooLarge;
        if (MinFileSizeKb > 0 && size < MinFileSizeKb * 1024L)
            return FilterVerdict.TooSmall;
        return FilterVerdict.Accepted;
    }

    internal IEnumerable<ValidationIssue> Validate()
    {
        foreach (string mask in Include.Concat(Exclude).Where(m => string.IsNullOrWhiteSpace(m)))
            yield return new(IssueSeverity.Error, L.F("filters: порожня маска шляху «{0}».", mask));
        if (ModifiedAfter is { } a && ModifiedBefore is { } b && a >= b)
            yield return new(IssueSeverity.Error, L.T("filters: modifiedAfter має бути раніше за modifiedBefore."));
        if (Extensions.Count == 0)
            yield return new(IssueSeverity.Error, L.T("filters.extensions: не задано жодного розширення."));
        foreach (var ext in Extensions.Where(e => !SupportedExtensions.Contains(e, StringComparer.OrdinalIgnoreCase)))
            yield return new(IssueSeverity.Error, L.F("filters.extensions: «{0}» не підтримується (лише {1}).", ext, string.Join(", ", SupportedExtensions)));
        if (MaxFileSizeMb < 0)
            yield return new(IssueSeverity.Error, L.T("filters.maxFileSizeMb не може бути від'ємним."));
        if (MinFileSizeKb < 0)
            yield return new(IssueSeverity.Error, L.T("filters.minFileSizeKb не може бути від'ємним."));
        if (MinFileSizeKb > 0 && MaxFileSizeMb > 0 && MinFileSizeKb * 1024L >= MaxFileSizeMb * 1024L * 1024L)
            yield return new(IssueSeverity.Error, L.T("filters: мінімальний розмір має бути меншим за максимальний."));
    }
}

public sealed class ProfileException(string message, Exception? inner = null) : Exception(message, inner);
