using SPOReplaceLink.Core.Rules;

namespace SPOReplaceLink.Core.Tests;

public class RuleSetTests
{
    private const string Host = "https://contoso.sharepoint.com";

    /// <summary>The rules from the ReplaceMagic screenshot, in their original order.</summary>
    private static readonly RewriteRule[] ReplaceMagicRules =
    [
        new("/Document/", "https://contoso.sharepoint.com/sites/finance/"),
        new("https://contoso.sharepoint.com/Document", "https://contoso.sharepoint.com/sites/finance"),
        new("/Document/Analytics/", "/sites/finance/analytics_migrated/"),
        new("/sites/finance/Analytics/", "/sites/finance/analytics_migrated/"),
    ];

    private static RuleSet Build(string? baseHost, params RewriteRule[] rules)
    {
        var build = RuleSet.Build(rules, baseHost);
        Assert.True(build.IsValid, string.Join("\n", build.Issues));
        return build.RuleSet!;
    }

    private static List<ValidationIssue> Errors(string? baseHost, params RewriteRule[] rules) =>
        RuleSet.Build(rules, baseHost).Issues.Where(i => i.Severity == IssueSeverity.Error).ToList();

    [Theory]
    [InlineData(Host + "/Document/Fin/a.xlsx", Host + "/sites/finance/Fin/a.xlsx")]
    [InlineData(Host + "/Document/Analytics/a.xlsx", Host + "/sites/finance/analytics_migrated/a.xlsx")]
    [InlineData(Host + "/sites/finance/Analytics/a.xlsx", Host + "/sites/finance/analytics_migrated/a.xlsx")]
    [InlineData("https://CONTOSO.sharepoint.com/DOCUMENT/analytics/a.xlsx", Host + "/sites/finance/analytics_migrated/a.xlsx")]
    [InlineData("/Document/Fin/a.xlsx", Host + "/sites/finance/Fin/a.xlsx")]
    [InlineData(Host + "/Document", Host + "/sites/finance")]
    [InlineData(Host + "/Document/", Host + "/sites/finance/")]
    [InlineData(Host + "/Document/Fin/a.xlsx#Sheet1!A1", Host + "/sites/finance/Fin/a.xlsx#Sheet1!A1")]
    public void ReplaceMagicRulesWorkInOnePass(string link, string expected)
    {
        var rules = Build(Host, ReplaceMagicRules);
        var result = rules.Rewrite(link);
        Assert.Equal(RewriteOutcome.Replaced, result.Outcome);
        Assert.Equal(expected, result.NewValue);
        // P-4: the result is final.
        Assert.False(rules.Rewrite(result.NewValue!).Changed);
    }

    [Fact]
    public void ReplaceMagicRulesValidateWithDuplicateWarningOnly()
    {
        var build = RuleSet.Build(ReplaceMagicRules, Host);
        Assert.True(build.IsValid);
        var issue = Assert.Single(build.Issues);
        Assert.Equal(IssueSeverity.Warning, issue.Severity);
        Assert.Equal(2, issue.RuleIndex);
    }

    [Fact]
    public void RuleOrderDoesNotMatter()
    {
        string[] links =
        [
            Host + "/Document/Analytics/Q1/a.xlsx", Host + "/Document/x.xlsx",
            Host + "/sites/finance/Analytics/b.xlsx", Host + "/Documents/c.xlsx",
        ];
        var expected = links.Select(l => Build(Host, ReplaceMagicRules).Rewrite(l).NewValue).ToList();
        var reversed = Build(Host, ReplaceMagicRules.Reverse().ToArray());
        Assert.Equal(expected, links.Select(l => reversed.Rewrite(l).NewValue));
    }

    [Theory]
    [InlineData(Host + "/Documents/a.xlsx")]
    [InlineData(Host + "/MyDocument/a.xlsx")]
    [InlineData(Host + "/Document_old/a.xlsx")]
    [InlineData("https://other.sharepoint.com/Document/a.xlsx")]
    public void MatchesOnlyAtSegmentBoundary(string link) =>
        Assert.Equal(RewriteOutcome.NoRule, Build(Host, ReplaceMagicRules).Rewrite(link).Outcome);

    [Fact]
    public void FullyEncodedLinkGetsFullyEncodedPrefix()
    {
        var rules = Build(null, new RewriteRule("https://h/Old Lib", "https://h/sites/Нова бібліотека"));
        Assert.Equal("https://h/sites/%D0%9D%D0%BE%D0%B2%D0%B0%20%D0%B1%D1%96%D0%B1%D0%BB%D1%96%D0%BE%D1%82%D0%B5%D0%BA%D0%B0/%D0%B0%20b.xlsx",
            rules.Rewrite("https://h/Old%20Lib/%D0%B0%20b.xlsx").NewValue);
    }

    [Fact]
    public void ExcelStyleLinkKeepsItsTailAndStyle()
    {
        // Excel writes spaces as %20 and Cyrillic as is.
        var rules = Build(null, new RewriteRule("https://h/Document", "https://h/sites/Нова бібліотека"));
        Assert.Equal("https://h/sites/Нова%20бібліотека/Fin/Звіт%202024.xlsx",
            rules.Rewrite("https://h/Document/Fin/Звіт%202024.xlsx").NewValue);
        Assert.Equal("https://h/sites/Нова%20бібліотека/a.xlsx", rules.Rewrite("https://h/Document/a.xlsx").NewValue);
    }

    [Fact]
    public void KeepsRawSpelling()
    {
        var rules = Build(null, new RewriteRule("https://h/Old Lib", "https://h/sites/Нова бібліотека"));
        Assert.Equal("https://h/sites/Нова бібліотека/a b.xlsx", rules.Rewrite("https://h/Old Lib/a b.xlsx").NewValue);
    }

    [Fact]
    public void UncToHttpsConvertsSeparatorsAndEncodesSpaces()
    {
        var rules = Build(null, new RewriteRule(@"\\fs01\Finance", "https://h/sites/finance/Shared Documents"));
        Assert.Equal("https://h/sites/finance/Shared%20Documents/Q1/Звіт%201.xlsx",
            rules.Rewrite(@"\\fs01\Finance\Q1\Звіт 1.xlsx").NewValue);
        Assert.Equal("https://h/sites/finance/Shared%20Documents/a.xlsx",
            rules.Rewrite(@"file:///\\FS01\finance\a.xlsx").NewValue);
    }

    [Fact]
    public void KeepsFileSchemeForFileSystemTargets()
    {
        var rules = Build(null, new RewriteRule(@"C:\Data\Old", @"\\fs02\New"));
        Assert.Equal(@"file:///\\fs02\New\a.xlsx", rules.Rewrite(@"file:///C:\Data\Old\a.xlsx").NewValue);
        Assert.Equal(@"\\fs02\New\a.xlsx", rules.Rewrite(@"C:\Data\Old\a.xlsx").NewValue);
    }

    [Fact]
    public void PathWithoutDriveResolvesAgainstDocumentDrive()
    {
        // Excel stores a link to a file on the same drive as "/Users/…".
        var rules = Build(Host, new RewriteRule(@"C:\Data\Old", @"D:\New"));
        var result = rules.Rewrite("/Data/Old/Q1/a.xlsx", @"C:\Data\Reports");
        Assert.Equal(@"D:\New\Q1\a.xlsx", result.NewValue);
        Assert.Equal("C:/Data/Old/Q1/a.xlsx", result.ResolvedFrom);
    }

    [Fact]
    public void HttpToFileSystemPathDecodes()
    {
        var rules = Build(null, new RewriteRule("https://h/Old", @"\\fs02\Archive"));
        Assert.Equal(@"\\fs02\Archive\Q1\a b.xlsx", rules.Rewrite("https://h/Old/Q1/a%20b.xlsx").NewValue);
    }

    [Fact]
    public void RelativeLinksResolveAgainstOldLocation()
    {
        var rules = Build(Host, ReplaceMagicRules);
        var result = rules.Rewrite("../Analytics/a.xlsx", Host + "/Document/Fin");
        Assert.Equal(Host + "/sites/finance/analytics_migrated/a.xlsx", result.NewValue);
        Assert.Equal(Host + "/Document/Analytics/a.xlsx", result.ResolvedFrom);

        Assert.Equal(RewriteOutcome.Unresolved, rules.Rewrite("../Analytics/a.xlsx").Outcome);
        Assert.Equal(RewriteOutcome.NoRule, rules.Rewrite("a.xlsx", "https://other/x").Outcome);
    }

    [Fact]
    public void ServerRelativeLinkWithoutBaseHostIsUnresolved()
    {
        var rules = Build(null, new RewriteRule("https://h/Old", "https://h/New"));
        Assert.Equal(RewriteOutcome.Unresolved, rules.Rewrite("/Old/a.xlsx").Outcome);
    }

    [Theory]
    [InlineData("#Sheet1!A1")]
    [InlineData("")]
    public void InternalLinksAreNotTouched(string link) =>
        Assert.Equal(RewriteOutcome.NotApplicable, Build(Host, ReplaceMagicRules).Rewrite(link).Outcome);

    [Fact]
    public void ConflictIsError()
    {
        var errors = Errors(null, new("https://h/A", "https://h/B"), new("https://h/a/", "https://h/C"));
        Assert.Contains(errors, e => e.Message.Contains("Конфлікт") && e.RuleIndex == 2);
    }

    [Fact]
    public void CycleIsError()
    {
        var errors = Errors(null, new("https://h/A", "https://h/B"), new("https://h/B", "https://h/A"));
        Assert.Equal(2, errors.Count);
    }

    [Fact]
    public void NewPrefixInsideOldIsError()
    {
        var errors = Errors(null, new RewriteRule("https://h/Lib", "https://h/Lib/Archive"));
        Assert.Contains(errors, e => e.Message.Contains("цим самим правилом"));
    }

    [Fact]
    public void OutputMatchingLongerRuleIsError()
    {
        // https://h/A/X/… → https://h/B/X/…, which rule 2 would rewrite again.
        var errors = Errors(null, new("https://h/A", "https://h/B"), new("https://h/B/X", "https://h/C"));
        var error = Assert.Single(errors);
        Assert.Equal(1, error.RuleIndex);
    }

    [Fact]
    public void OutputMatchingLongerRuleIsFineWhenCoveredBySpecificRule()
    {
        var build = RuleSet.Build(
            [new("https://h/A", "https://h/B"), new("https://h/B/X", "https://h/C"), new("https://h/A/X", "https://h/C")], null);
        Assert.True(build.IsValid, string.Join("\n", build.Issues));
    }

    [Fact]
    public void IdentityRuleIsIgnoredWithWarning()
    {
        var build = RuleSet.Build([new("https://h/A", "https://H/a/"), new("https://h/X", "https://h/Y")], null);
        Assert.True(build.IsValid);
        Assert.Single(build.RuleSet!.Rules);
        Assert.Contains(build.Issues, i => i.Severity == IssueSeverity.Warning && i.RuleIndex == 1);
    }

    [Theory]
    [InlineData("relative/path", "https://h/x")]
    [InlineData("https://h/x", "")]
    [InlineData("/no/host", "https://h/x")]
    public void InvalidPrefixesAreErrors(string old, string @new) =>
        Assert.NotEmpty(Errors(null, new RewriteRule(old, @new)));

    [Fact]
    public void BaseHostWithPathIsError() =>
        Assert.NotEmpty(Errors("https://h/sites/x", new RewriteRule("https://h/A", "https://h/B")));

    [Fact]
    public void EmptyRuleSetIsError() => Assert.NotEmpty(Errors(null));

    [Fact]
    public void RewriteIsIdempotentForManyLinks()
    {
        var rules = Build(Host, ReplaceMagicRules);
        string[] tails = ["", "/", "/a.xlsx", "/Analytics", "/Analytics/b c.xlsx", "/analytics_migrated/x", "/Fin/Звіт%201.xlsx", "?web=1"];
        string[] heads = [Host + "/Document", Host + "/sites/finance", Host + "/sites/finance/Analytics", "/Document", "/sites/finance/Analytics"];
        foreach (string head in heads)
        {
            foreach (string tail in tails)
            {
                var once = rules.Rewrite(head + tail);
                string value = once.NewValue ?? head + tail;
                Assert.False(rules.Rewrite(value).Changed, $"{head + tail} → {value} змінюється вдруге");
            }
        }
    }
}
