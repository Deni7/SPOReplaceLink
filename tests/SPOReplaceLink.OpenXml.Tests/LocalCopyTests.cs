using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Rules;

namespace SPOReplaceLink.OpenXml.Tests;

/// <summary>Links Excel rewrote relative to the temporary folder of an intermediate copy (ExcelWorker, stage 4).</summary>
public class LocalCopyTests
{
    private static readonly LocalCopy Copy = new(@"C:\T\run-1\x1\1\2\3\4\5\6\7\8", @"C:\T\run-1\x1");

    [Theory]
    // Was absolute on the same drive: Excel wrote it relative to the temp folder, or without the drive, or as file:///.
    [InlineData("../../../../../../../../../../../Old/Document/a.xlsx", LinkKind.LinkSource, @"file:///C:\Old\Document\a.xlsx")]
    [InlineData("/Old/Document/a.xlsx", LinkKind.LinkSource, @"file:///C:\Old\Document\a.xlsx")]
    [InlineData("../../../../../../../../../../../Old/Document/Звіт%201.xlsx", LinkKind.Hyperlink, @"file:///C:\Old\Document\Звіт 1.xlsx")]
    // Was relative to the document: stays relative, resolved against the old folder as usual.
    [InlineData("../Analytics/a.xlsx", LinkKind.Hyperlink, null)]
    [InlineData(@"file:///C:\T\run-1\x1\1\2\3\4\5\6\7\Analytics\a.xlsx", LinkKind.LinkSource, "../Analytics/a.xlsx")]
    [InlineData("/T/run-1/x1/1/2/3/4/5/6/7/Analytics/a.xlsx", LinkKind.LinkSource, "../Analytics/a.xlsx")]
    // Other forms and places are left alone.
    [InlineData("https://h/Document/a.xlsx", LinkKind.Hyperlink, null)]
    [InlineData(@"\\server\share\a.xlsx", LinkKind.LinkSource, null)]
    [InlineData("/Document/a.xlsx", LinkKind.Hyperlink, null)]
    [InlineData("../a.xlsx", LinkKind.Formula, null)]
    public void Translate(string value, LinkKind kind, string? expected) => Assert.Equal(expected, Copy.Translate(value, kind));

    [Fact]
    public void CreateMakesTheFolderBelowTheRoot()
    {
        string root = Path.Combine(Path.GetTempPath(), "spo-rl-local", Guid.NewGuid().ToString("N"));
        var copy = LocalCopy.Create(root);
        Assert.True(Directory.Exists(copy.Folder));
        Assert.Equal(LocalCopy.Depth, Path.GetRelativePath(copy.Root, copy.Folder).Split(Path.DirectorySeparatorChar).Length);
        Directory.Delete(root, true);
    }

    [Fact]
    public void ServerRelativeLinkSourceStillMatchesTheUsualRule()
    {
        // "/Document/…" written for a SharePoint document: the local reading matches no rule, the usual one does.
        var rules = RuleSet.Build([new("/Document/", TestWorkbook.Host + "/sites/finance/")], TestWorkbook.Host).RuleSet!;
        string root = Path.Combine(Path.GetTempPath(), "spo-rl-local", Guid.NewGuid().ToString("N"));
        var copy = LocalCopy.Create(root);
        var residuals = new WorkbookProcessor(rules, new PlacesOptions())
            .CheckLinks([("excel:link_sources", LinkKind.LinkSource, "/Document/Fin/a.xlsx")], TestWorkbook.Host + "/Document/Fin", copy);
        Assert.Equal("/Document/Fin/a.xlsx", Assert.Single(residuals).Match);
        Directory.Delete(root, true);
    }
}
