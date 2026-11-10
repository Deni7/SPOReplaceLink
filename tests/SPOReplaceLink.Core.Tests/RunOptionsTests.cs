using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Rules;

namespace SPOReplaceLink.Core.Tests;

public class RunOptionsTests
{
    private static Connection Conn(string name) => new()
    {
        Name = name,
        TenantId = "t",
        ClientId = Guid.NewGuid().ToString(),
        CertificateThumbprint = new string('B', 40),
    };

    private static Profile Valid() => new()
    {
        Rules = [new RuleEntry { Old = "https://h/a", New = "https://h/b" }],
        Targets = [new Target { SiteUrl = "https://h/sites/x", Library = "Documents" }],
        Connections = ["c1"],
    };

    [Fact]
    public void ValidProfileForRun() => Assert.True(Valid().ValidateForRun([Conn("c1")]).IsValid);

    [Theory]
    [InlineData(0, 5, true)]
    [InlineData(16, 10, true)]
    [InlineData(0, 120, true)]
    [InlineData(-1, 5, false)]
    [InlineData(17, 5, false)]
    [InlineData(2, 0, false)]
    public void ExcelProcessesAndTimeoutAreChecked(int processes, int timeout, bool valid)
    {
        var p = Valid();
        p.Performance.ExcelProcesses = processes;
        p.Performance.ExcelTimeoutMinutes = timeout;
        Assert.Equal(valid, p.ValidateForRun([Conn("c1")]).IsValid);
    }

    [Fact]
    public void ExcelTimeoutLongerThanTheFileTimeoutIsAWarning()
    {
        var p = Valid();
        p.Performance.ExcelTimeoutMinutes = 11;
        var v = p.ValidateForRun([Conn("c1")]);
        Assert.True(v.IsValid);
        Assert.Contains(v.Issues, i => i.Severity == IssueSeverity.Warning && i.Message.Contains("Excel"));
    }
    [Fact]
    public void SeveralConnectionsNeedPool()
    {
        var p = Valid();
        p.Connections = ["c1", "c2"];
        var v = p.ValidateForRun([Conn("c1"), Conn("c2")]);
        Assert.False(v.IsValid);
        p.Pool.Enabled = true;
        v = p.ValidateForRun([Conn("c1"), Conn("c2")]);
        Assert.True(v.IsValid);
        Assert.Contains(v.Issues, i => i.Severity == IssueSeverity.Warning && i.Message.Contains("пул"));
    }

    [Fact]
    public void MissingTargetsAndConnectionsAreErrors()
    {
        var p = Valid();
        p.Targets.Clear();
        p.Connections = ["nope"];
        var v = p.ValidateForRun([]);
        Assert.Contains(v.Issues, i => i.Message.Contains("targets"));
        Assert.Contains(v.Issues, i => i.Message.Contains("nope"));
    }

    [Fact]
    public void ConnectionValidation()
    {
        var c = Conn("c");
        Assert.Empty(c.Validate());
        c.CertificateThumbprint = "short";
        c.ClientId = "x";
        Assert.Equal(2, c.Validate().Count());
    }

    [Fact]
    public void TemplatesAreResolvedBeforeInlineRules()
    {
        var p = Valid();
        p.Templates = ["fin"];
        var templates = new Dictionary<string, Template>(StringComparer.OrdinalIgnoreCase)
        {
            ["fin"] = new Template { Name = "fin", Rules = [new RuleEntry { Old = "https://h/x", New = "https://h/y" }] },
        };
        var resolved = p.ResolveTemplates(templates);
        Assert.Empty(resolved.Templates);
        Assert.Equal(["https://h/x", "https://h/a"], resolved.Rules.Select(r => r.Old));
        Assert.Throws<ProfileException>(() => new Profile { Templates = ["missing"] }.ResolveTemplates(templates));
        Assert.False(p.Validate().IsValid); // unresolved templates are not for local use
    }

    [Theory]
    [InlineData("Mon", "20:00", "07:00", "2026-10-26 21:00", true)] // Monday evening
    [InlineData("Mon", "20:00", "07:00", "2026-10-27 06:59", true)] // Tuesday morning belongs to Monday's window
    [InlineData("Mon", "20:00", "07:00", "2026-10-27 21:00", false)]
    [InlineData("", "09:00", "18:00", "2026-10-25 10:00", true)]
    [InlineData("Sat,Sun", "00:00", "24:00", "2026-10-26 10:00", false)]
    public void WindowContains(string days, string from, string to, string at, bool expected)
    {
        var w = new ScheduleWindow { Days = days.Split(',', StringSplitOptions.RemoveEmptyEntries).ToList(), From = from, To = to };
        Assert.Equal(expected, w.Contains(DateTime.Parse(at, System.Globalization.CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void BudgetFollowsWindows()
    {
        var perf = new PerformanceOptions
        {
            BudgetPercent = 50,
            TimeZone = "UTC",
            Windows = [new ScheduleWindow { From = "20:00", To = "07:00", BudgetPercent = 90 }],
        };
        Assert.Equal(90, perf.BudgetAt(new DateTimeOffset(2026, 10, 23, 22, 0, 0, TimeSpan.Zero)));
        Assert.Equal(50, perf.BudgetAt(new DateTimeOffset(2026, 10, 23, 12, 0, 0, TimeSpan.Zero)));
        perf.OnlyInWindows = true;
        Assert.Null(perf.BudgetAt(new DateTimeOffset(2026, 10, 23, 12, 0, 0, TimeSpan.Zero)));
    }

    [Theory]
    [InlineData("Fin/**", "Fin/2024/a.xlsx", true)]
    [InlineData("**/Archive/**", "Fin/Archive/a.xlsx", true)]
    [InlineData("**/Archive/**", "Archive/a.xlsx", true)]
    [InlineData("*.xlsx", "a.xlsx", true)]
    [InlineData("*.xlsx", "Fin/a.xlsx", false)]
    [InlineData("Fin/?.xlsx", "fin/b.xlsx", true)]
    public void PathMasks(string mask, string path, bool expected) => Assert.Equal(expected, PathMask.IsMatch(mask, path));

    [Fact]
    public void FiltersCheckMinimumSize()
    {
        var f = new FileFilters { MinFileSizeKb = 10 };
        Assert.Equal(FilterVerdict.TooSmall, f.Check("a.xlsx", 9 * 1024, null));
        Assert.Equal(FilterVerdict.Accepted, f.Check("a.xlsx", 10 * 1024, null));
        f.MaxFileSizeMb = 0;
        Assert.Equal(FilterVerdict.Accepted, f.Check("a.xlsx", 900L * 1024 * 1024, null));

        var p = Valid();
        p.Filters.MinFileSizeKb = 2048;
        p.Filters.MaxFileSizeMb = 1;
        Assert.Contains(p.ValidateForRun([Conn("c1")]).Issues, i => i.Message.Contains("мінімальний розмір"));
    }

    [Fact]
    public void FiltersCheckDateAndSize()
    {
        var f = new FileFilters { ModifiedAfter = new DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero), MaxFileSizeMb = 1 };
        Assert.Equal(FilterVerdict.Date, f.Check("a.xlsx", 10, new DateTimeOffset(2023, 5, 1, 0, 0, 0, TimeSpan.Zero)));
        Assert.Equal(FilterVerdict.TooLarge, f.Check("a.xlsx", 2 << 20, DateTimeOffset.UtcNow));
        Assert.Equal(FilterVerdict.Accepted, f.Check("a.xlsx", 10, DateTimeOffset.UtcNow));
        Assert.Equal(FilterVerdict.Extension, f.Check("a.docx", 10, DateTimeOffset.UtcNow));
    }

    [Fact]
    public void LimitsByTier()
    {
        var c = Conn("c");
        c.Licenses = LicenseTier.UpTo5000;
        Assert.Equal((2_500, 2_400_000, 37_500), c.Limits());
        c.AppPerMinuteOverride = 1000;
        Assert.Equal(1000, c.Limits().PerMinute);
    }

    [Fact]
    public void ReverseRuleGivesOldLocation()
    {
        var rules = RuleSet.Build([new RewriteRule("https://h/Document", "https://h/sites/finance/Shared Documents")], null).RuleSet!;
        Assert.Equal("https://h/Document/Fin", rules.Reverse("https://h/sites/finance/Shared%20Documents/Fin"));
        Assert.Null(rules.Reverse("https://h/elsewhere"));

        var ambiguous = RuleSet.Build([new RewriteRule("https://h/A", "https://h/N"), new RewriteRule("https://h/B", "https://h/N/")], null);
        Assert.Null(ambiguous.RuleSet?.Reverse("https://h/N/x"));
    }

    [Fact]
    public void TemplateKeepsItsTestHostAndGuessesOneFromAbsoluteRules()
    {
        var plain = new Template { Name = "t", Rules = [new RuleEntry { Old = "/Document/A/", New = "https://contoso.sharepoint.com/sites/x/A/" }] };
        Assert.Equal("https://contoso.sharepoint.com", plain.GuessBaseHost());
        Assert.DoesNotContain("testBaseHost", plain.ToJson());
        Assert.Null(new Template { Rules = [new RuleEntry { Old = "/a/", New = "/b/" }] }.GuessBaseHost());

        plain.TestBaseHost = "https://other.sharepoint.com";
        var back = Template.Parse(plain.ToJson());
        Assert.Equal("https://other.sharepoint.com", back.TestBaseHost);
        Assert.Equal("https://other.sharepoint.com", back.GuessBaseHost());
    }
}