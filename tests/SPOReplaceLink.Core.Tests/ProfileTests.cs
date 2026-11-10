using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Reports;
using SPOReplaceLink.Core.Rules;

namespace SPOReplaceLink.Core.Tests;

public class ProfileTests
{
    [Fact]
    public void SampleProfileIsValid()
    {
        string path = Path.Combine(FindRepoRoot(), "samples", "finance.profile.json");
        var validation = Profile.Load(path).Validate();
        Assert.True(validation.IsValid, string.Join("\n", validation.Issues));
        Assert.Equal(3, validation.RuleSet!.Rules.Count);
    }

    [Fact]
    public void UnknownPropertyIsRejected()
    {
        var ex = Assert.Throws<ProfileException>(() => Profile.Parse("""{ "rules": [], "place": {} }"""));
        Assert.Contains("place", ex.Message);
    }

    [Fact]
    public void UnsupportedPlaceIsError()
    {
        var profile = Profile.Parse("""
            { "rules": [ { "old": "https://h/a", "new": "https://h/b" } ],
              "places": { "powerQuery": true } }
            """);
        var validation = profile.Validate();
        Assert.False(validation.IsValid);
        Assert.Contains(validation.Issues, i => i.Message.Contains("powerQuery"));
    }

    [Fact]
    public void UnsupportedExtensionIsError()
    {
        var profile = Profile.Parse("""
            { "rules": [ { "old": "https://h/a", "new": "https://h/b" } ],
              "filters": { "extensions": [ ".docx" ] } }
            """);
        Assert.False(profile.Validate().IsValid);
    }

    [Fact]
    public void RoundTripsThroughJson()
    {
        var profile = Profile.Parse("""
            // comment
            { "name": "x", "baseHost": "https://h", "rules": [ { "old": "/a", "new": "/b", "name": "n" }, ] }
            """);
        var again = Profile.Parse(profile.ToJson());
        Assert.Equal("x", again.Name);
        Assert.Equal("/b", again.Rules[0].New);
        Assert.True(again.Places.Hyperlinks);
    }

    [Theory]
    [InlineData("a.xlsx", true)]
    [InlineData("A.XLSB", true)]
    [InlineData("~$a.xlsx", false)]
    [InlineData("a.docx", false)]
    public void FiltersByExtension(string name, bool accepted) =>
        Assert.Equal(accepted, new FileFilters().Accepts(Path.Combine("dir", name)));

    [Theory]
    [InlineData("plain", "plain")]
    [InlineData("a,b", "\"a,b\"")]
    [InlineData("say \"hi\"", "\"say \"\"hi\"\"\"")]
    [InlineData("=HYPERLINK(1)", "'=HYPERLINK(1)")]
    [InlineData(null, "")]
    public void EscapesCsv(string? value, string expected) => Assert.Equal(expected, CsvReport.Escape(value, ','));

    [Fact]
    public void WritesReports()
    {
        string dir = Path.Combine(Path.GetTempPath(), "spo-rl-tests", Guid.NewGuid().ToString("N"));
        var result = new FileResult { Path = @"Fin\a.xlsx", Format = "xlsx", Status = FileStatus.Done };
        result.Links.Add(new LinkRecord("xl/worksheets/_rels/sheet1.xml.rels", LinkKind.Hyperlink, "https://h/a", "https://h/b", LinkAction.Replaced, 1, null));
        result.Residuals.Add(new Residual("xl/sharedStrings.xml", "utf-8", "https://h/a", "…https://h/a…"));
        using (var report = new CsvReport(dir))
            report.Add(result);

        string[] files = File.ReadAllLines(Path.Combine(dir, "files.csv"));
        Assert.Equal(2, files.Length);
        Assert.StartsWith(@"Fin\a.xlsx,xlsx,done:residual,1,1,", files[1]);
        Assert.Contains("hyperlink,replaced,https://h/a,https://h/b,1", File.ReadAllText(Path.Combine(dir, "links.csv")));
        Assert.Equal(2, File.ReadAllLines(Path.Combine(dir, "residuals.csv")).Length);
        Directory.Delete(dir, true);
    }

    [Fact]
    public void StatusCodes()
    {
        var r = new FileResult { Path = "a", Format = "xlsb", Status = FileStatus.Planned, NeedsExcel = true };
        Assert.Equal("planned:needs_excel", r.StatusCode);
        r = new FileResult { Path = "a", Format = "xlsx", Status = FileStatus.Failed, ErrorClass = ErrorClass.Validation };
        Assert.Equal("failed:validation", r.StatusCode);
        Assert.True(r.NeedsAttention);
    }

    internal static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "SPOReplaceLink.sln")))
            dir = dir.Parent;
        return dir?.FullName ?? throw new InvalidOperationException("Корінь репозиторію не знайдено.");
    }
}
