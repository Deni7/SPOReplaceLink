using SPOReplaceLink.Core.Processing;

namespace SPOReplaceLink.Engine.Tests;

/// <summary>The prefixes report (inventory): where a prefix is cut and the depth setting.</summary>
public class PrefixReportTests
{
    [Theory]
    [InlineData("https://h.sharepoint.com/sites/fin/Shared%20Documents/a.xlsx", 1, "https://h.sharepoint.com/sites")]
    [InlineData("https://h.sharepoint.com/sites/fin/Shared%20Documents/a.xlsx", 2, "https://h.sharepoint.com/sites/fin")]
    [InlineData("https://h.sharepoint.com/sites/fin/Shared%20Documents/a.xlsx", 3, "https://h.sharepoint.com/sites/fin/Shared Documents")]
    [InlineData("https://h.sharepoint.com/sites/fin?web=1", 5, "https://h.sharepoint.com/sites/fin")]
    [InlineData(@"\\fs01\share\Fin\Q1\a.xlsx", 1, @"\\fs01\share")]
    [InlineData(@"\\fs01\share\Fin\Q1\a.xlsx", 3, @"\\fs01\share\Fin\Q1")]
    [InlineData(@"C:\Users\ivan\Documents\a.xlsx", 2, @"C:\Users\ivan")]
    [InlineData(@"C:\Users\ivan\Documents\a.xlsx", 3, @"C:\Users\ivan\Documents")]
    [InlineData("/sites/fin/Shared Documents/a.xlsx", 1, "/sites")]
    [InlineData("/sites/fin/Shared Documents/a.xlsx", 3, "/sites/fin/Shared Documents")]
    public void PrefixKeepsRootAndDepthSegments(string link, int depth, string expected) =>
        Assert.Equal(expected, ReportExporter.PrefixOf(link, depth));

    [Fact]
    public void DefaultDepthIsTwo() =>
        Assert.Equal(@"\\fs01\share\Fin", ReportExporter.PrefixOf(@"\\fs01\share\Fin\Q1\a.xlsx"));

    [Fact]
    public void DeeperPrefixesSplitTheHistogram()
    {
        var links = new[]
        {
            ("a.xlsx", Link(@"\\fs01\share\Fin\a.xlsx", LinkAction.NoRule)),
            ("b.xlsx", Link(@"\\fs01\share\HR\b.xlsx", LinkAction.Replaced)),
        };
        var shallow = ReportExporter.Prefixes(links, 1);
        var row = Assert.Single(shallow);
        Assert.Equal((2L, 2L, 1L, 1L), (row.Links, row.Files, row.Replaced, row.NoRule));

        var deep = ReportExporter.Prefixes(links, 2);
        Assert.Equal([@"\\fs01\share\Fin", @"\\fs01\share\HR"], deep.Select(p => p.Prefix).Order(StringComparer.Ordinal));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(ReportExporter.MaxPrefixDepth + 1)]
    public void DepthOutOfRangeIsRejected(int depth) =>
        Assert.Throws<ArgumentOutOfRangeException>(() => ReportExporter.Prefixes([], depth));

    private static LinkRecord Link(string value, LinkAction action) =>
        new("xl/externalLinks/externalLink1.xml", LinkKind.LinkSource, value, null, action, null, null);
}
