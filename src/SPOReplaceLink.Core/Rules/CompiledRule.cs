using SPOReplaceLink.Core.Urls;

namespace SPOReplaceLink.Core.Rules;

/// <summary>A rule resolved to absolute prefixes and prepared for matching.</summary>
public sealed class CompiledRule
{
    internal CompiledRule(int index, RewriteRule source, string effectiveOld, string effectiveNew)
    {
        Index = index;
        Source = source;
        EffectiveOld = effectiveOld;
        EffectiveNew = effectiveNew;
        MatchKey = TrimTrailingSlash(NormalizedLink.Normalize(effectiveOld));
        NewKey = TrimTrailingSlash(NormalizedLink.Normalize(effectiveNew));
        NewForm = LinkClassifier.Classify(effectiveNew);
        NewCanonical = TrimTrailingSlash(UrlText.Decode(effectiveNew));
    }

    /// <summary>1-based position of the rule in the profile, for messages and reports.</summary>
    public int Index { get; }

    public RewriteRule Source { get; }

    /// <summary>Old prefix as an absolute location (server-relative rules get the base host).</summary>
    public string EffectiveOld { get; }

    public string EffectiveNew { get; }

    /// <summary>Normalized old prefix without trailing slash; links are compared against it.</summary>
    public string MatchKey { get; }

    /// <summary>Normalized new prefix without trailing slash.</summary>
    public string NewKey { get; }

    public LinkForm NewForm { get; }

    /// <summary>New prefix decoded and without trailing slash; encoded again on output if needed.</summary>
    public string NewCanonical { get; }

    public override string ToString() => $"#{Index} {Source.Old} → {Source.New}";

    internal static string TrimTrailingSlash(string s)
    {
        int end = s.Length;
        while (end > 0 && LinkClassifier.IsSlash(s[end - 1]) && !EndsAtRoot(s, end))
            end--;
        return s[..end];
    }

    // Keep "//" of a UNC root and "C:/" of a drive root.
    private static bool EndsAtRoot(string s, int end) =>
        end <= 2 || (end == 3 && s[1] == ':');
}
