using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Rules;

namespace SPOReplaceLink.OpenXml.Parts;

/// <summary>Applies the rule set to links found in one part and records what happened to each.</summary>
internal sealed class LinkContext(RuleSet rules, PlacesOptions places, string? documentFolder, string part, LocalCopy? local = null)
{
    public string Part { get; } = part;

    public List<LinkRecord> Links { get; } = [];

    public int Changes { get; private set; }

    /// <summary>Returns the replacement, or null if the link stays as it is.</summary>
    public string? Rewrite(string value, LinkKind kind)
    {
        string subject = value;
        var result = rules.Rewrite(value, documentFolder);
        // An intermediate copy from Excel: the link as it was before Excel rewrote it relative to the temporary folder.
        if (local?.Translate(value, kind) is { } original)
        {
            var translated = rules.Rewrite(original, documentFolder);
            if (translated.Outcome is RewriteOutcome.Replaced or RewriteOutcome.AlreadyCurrent || result.Outcome is not (RewriteOutcome.Replaced or RewriteOutcome.AlreadyCurrent))
            {
                subject = original;
                result = translated;
            }
        }
        if (result.Outcome == RewriteOutcome.NotApplicable)
            return null;

        var action = result.Outcome switch
        {
            RewriteOutcome.Replaced => IsEnabled(kind) ? LinkAction.Replaced : LinkAction.PlaceDisabled,
            RewriteOutcome.AlreadyCurrent => LinkAction.Current,
            RewriteOutcome.Unresolved => LinkAction.Unresolved,
            _ => LinkAction.NoRule,
        };
        string? newValue = action == LinkAction.Replaced ? result.NewValue : null;
        Links.Add(new LinkRecord(Part, kind, subject, newValue, action, result.Rule?.Index, result.ResolvedFrom));
        if (newValue is null)
            return null;
        Changes++;
        return newValue;
    }

    private bool IsEnabled(LinkKind kind) => kind switch
    {
        LinkKind.Hyperlink => places.Hyperlinks,
        LinkKind.LinkSource => places.LinkSources,
        LinkKind.Formula or LinkKind.CachedValue => places.Formulas,
        LinkKind.WorkbookPath => true,
        _ => places.OtherExternal,
    };
}
