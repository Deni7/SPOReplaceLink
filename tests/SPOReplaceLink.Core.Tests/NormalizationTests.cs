using SPOReplaceLink.Core.Urls;

namespace SPOReplaceLink.Core.Tests;

public class NormalizationTests
{
    [Theory]
    [InlineData("https://CONTOSO.SharePoint.com/Document/A.xlsx", "https://contoso.sharepoint.com/document/a.xlsx")]
    [InlineData("https://h/Document/My%20File.xlsx", "https://h/document/my file.xlsx")]
    [InlineData("https://h/%D0%97%D0%B2%D1%96%D1%82.xlsx", "https://h/звіт.xlsx")]
    [InlineData("https://h/Звіт.xlsx", "https://h/звіт.xlsx")]
    [InlineData(@"\\fs01\Finance\A.xlsx", "//fs01/finance/a.xlsx")]
    [InlineData(@"file:///\\fs01\Finance\A.xlsx", "//fs01/finance/a.xlsx")]
    [InlineData("file://fs01/Finance/A.xlsx", "//fs01/finance/a.xlsx")]
    [InlineData("file:////fs01/Finance/A.xlsx", "//fs01/finance/a.xlsx")]
    [InlineData(@"file:///C:\Users\A.xlsx", "c:/users/a.xlsx")]
    [InlineData(@"C:\Users\A.xlsx", "c:/users/a.xlsx")]
    [InlineData("https://h/100%zz", "https://h/100%zz")]
    [InlineData("https://h/a%FFb", "https://h/a%ffb")]
    public void Normalizes(string input, string expected) =>
        Assert.Equal(expected, NormalizedLink.Normalize(input));

    [Fact]
    public void MapsBackToOriginal()
    {
        var n = NormalizedLink.Create("https://h/%D0%97/Tail%20X");
        int slash = n.Value.IndexOf("/tail", StringComparison.Ordinal);
        Assert.Equal("/Tail%20X", n.Original[n.OriginalIndex(slash)..]);
        Assert.Equal(n.Original.Length, n.OriginalIndex(n.Value.Length));
    }

    [Theory]
    [InlineData("https://h/a", LinkForm.Http)]
    [InlineData("HTTP://h/a", LinkForm.Http)]
    [InlineData(@"\\srv\share", LinkForm.Unc)]
    [InlineData("//srv/share", LinkForm.Unc)]
    [InlineData(@"D:\x", LinkForm.Drive)]
    [InlineData("file:///D:/x", LinkForm.File)]
    [InlineData("/sites/x", LinkForm.ServerRelative)]
    [InlineData("../x.xlsx", LinkForm.Relative)]
    [InlineData(@"Folder\x.xlsx", LinkForm.Relative)]
    [InlineData("#Sheet1!A1", LinkForm.Internal)]
    [InlineData("mailto:a@b.c", LinkForm.OtherScheme)]
    [InlineData("", LinkForm.Empty)]
    public void Classifies(string link, LinkForm form) => Assert.Equal(form, LinkClassifier.Classify(link));

    [Theory]
    [InlineData("https://h/Document/Fin", "../Analytics/a.xlsx", "https://h/Document/Analytics/a.xlsx")]
    [InlineData("https://h/Document/Fin", "b.xlsx", "https://h/Document/Fin/b.xlsx")]
    [InlineData("https://h/Document/Fin/", @"Sub\c.xlsx", "https://h/Document/Fin/Sub/c.xlsx")]
    [InlineData("https://h/Document/Fin", "./d.xlsx", "https://h/Document/Fin/d.xlsx")]
    [InlineData(@"\\fs01\Finance\Q1", @"..\Q2\e.xlsx", @"\\fs01\Finance\Q2\e.xlsx")]
    [InlineData(@"C:\Data\Q1", "f.xlsx", @"C:\Data\Q1\f.xlsx")]
    public void ResolvesRelative(string folder, string relative, string expected) =>
        Assert.Equal(expected, RelativeLinkResolver.Resolve(folder, relative));

    [Fact]
    public void RelativeAboveRootIsUnresolved() =>
        Assert.Null(RelativeLinkResolver.Resolve("https://h/Document", "../../x.xlsx"));

    [Theory]
    [InlineData("a b", "a%20b")]
    [InlineData("Звіт", "%D0%97%D0%B2%D1%96%D1%82")]
    [InlineData("a%20b c", "a%20b%20c")]
    [InlineData("x?y=1&z#f", "x?y=1&z#f")]
    public void EncodesOnlyUnsafe(string input, string expected) => Assert.Equal(expected, UrlText.EncodeUnsafe(input));
}
