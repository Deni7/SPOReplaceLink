namespace SPOReplaceLink.Core.Rules;

public enum RewriteOutcome
{
    /// <summary>A rule matched and the link changes.</summary>
    Replaced,
    /// <summary>No rule matched the link.</summary>
    NoRule,
    /// <summary>A rule matched but produced the same text.</summary>
    AlreadyCurrent,
    /// <summary>Relative or server-relative link with no base to resolve it against.</summary>
    Unresolved,
    /// <summary>Internal (#Sheet!A1), empty, or non-location link — never touched.</summary>
    NotApplicable,
}

public sealed record RewriteResult(
    RewriteOutcome Outcome,
    string Original,
    string? NewValue = null,
    CompiledRule? Rule = null,
    string? ResolvedFrom = null)
{
    public bool Changed => Outcome == RewriteOutcome.Replaced;
}
