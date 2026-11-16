using System.Diagnostics;
using System.Text;
using SPOReplaceLink.Core.Rules;
using static SPOReplaceLink.OpenXml.Tests.TestWorkbook;

namespace SPOReplaceLink.OpenXml.Tests;

public sealed class ResidualScannerTests
{
    private static RuleSet ManyRules(int count)
    {
        var rules = new List<RewriteRule>();
        for (int i = 0; i < count; i++)
            rules.Add(new($"/sites/dept{i:D3}/Shared Documents/Звіти {i}/", Host + $"/sites/new{i:D3}/Docs/"));
        var build = RuleSet.Build(rules, Host);
        Assert.True(build.IsValid, string.Join("\n", build.Issues));
        return build.RuleSet!;
    }

    [Fact]
    public void FindsEverySpellingOfEveryRule()
    {
        var scanner = new ResidualScanner(ManyRules(100));
        string text =
            $"a {Host}/sites/dept000/Shared Documents/Звіти 0/x.xlsx " +
            "b HTTPS://CONTOSO.SHAREPOINT.COM\\SITES\\DEPT042\\SHARED DOCUMENTS\\ЗВІТИ 42\\y.xlsx " +
            "c https://contoso.sharepoint.com%2Fsites%2Fdept077%2FShared%20Documents%2F%D0%B7%D0%B2%D1%96%D1%82%D0%B8%2077 " +
            "d /sites/dept099/Shared+Documents/Звіти 99/z " +
            "e /x/sites/dept001/Shared Documents/Звіти 1/ " +          // inside another path: not a residual
            "f /sites/dept002/Shared Documents/Звіти 2x/ " +           // longer name: not a residual
            "g /sites/dept003/Shared Documents/Звіти 3%2Fq";

        var found = scanner.ScanText("p", text).Select(r => r.Match).ToList();

        Assert.True(found.Count == 5, string.Join(" | ", found));
        Assert.Contains(found, m => m.Contains("dept000"));
        Assert.Contains(found, m => m.Contains("DEPT042"));
        Assert.Contains(found, m => m.Contains("dept077"));
        Assert.Contains(found, m => m.Contains("dept099"));
        Assert.Contains(found, m => m.Contains("dept003"));
    }

    [Fact]
    public void ManyRulesOnLargePartFinishQuickly()
    {
        var scanner = new ResidualScanner(ManyRules(100));
        var sb = new StringBuilder();
        int row = 0;
        while (sb.Length < 30_000_000)
        {
            row++;
            sb.Append($"<row r=\"{row}\"><c r=\"A{row}\" t=\"s\"><v>{row}</v></c><c r=\"B{row}\"><f>%2F/sites/dept{row % 100:D3}/Shared Documents/</f></c></row>");
        }
        sb.Append($"<v>{Host}/sites/dept050/Shared Documents/Звіти 50/</v>");

        var watch = Stopwatch.StartNew();
        var found = scanner.ScanText("xl/worksheets/sheet1.xml", sb.ToString()).ToList();
        watch.Stop();

        Assert.Single(found);
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(5), $"scan took {watch.Elapsed}");
    }
}
