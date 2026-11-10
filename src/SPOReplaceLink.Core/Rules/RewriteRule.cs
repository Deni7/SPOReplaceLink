using SPOReplaceLink.Core.Localization;
namespace SPOReplaceLink.Core.Rules;

/// <summary>Old location prefix → new location prefix (RUL-1).</summary>
public sealed record RewriteRule(string Old, string New, string? Name = null)
{
    public override string ToString() => Name is { Length: > 0 } ? $"{Name}: {Old} → {New}" : $"{Old} → {New}";
}

public enum IssueSeverity
{
    Warning,
    Error,
}

/// <summary>A problem found when a rule set or a profile is validated (RUL-8, CFG-7).</summary>
/// <param name="RuleIndex">1-based number of the rule as the user sees it, or null for the set as a whole.</param>
public sealed record ValidationIssue(IssueSeverity Severity, string Message, int? RuleIndex = null)
{
    public override string ToString()
    {
        string level = Severity == IssueSeverity.Error ? L.T("ПОМИЛКА") : L.T("Попередження");
        return RuleIndex is int i ? L.F("{0} [правило {1}]: {2}", level, i, Message) : $"{level}: {Message}";
    }
}
