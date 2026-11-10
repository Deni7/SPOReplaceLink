using System.Text.RegularExpressions;
using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.Ui.Services;

namespace SPOReplaceLink.Host.Tests;

/// <summary>
/// The English user guide (docs/USER_GUIDE.en.md) as help: the same sections as the Ukrainian one, so the "?" buttons and
/// links of the app lead to the same place in both languages.
/// </summary>
public sealed partial class GuideTranslationTests
{
    private static string Root()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "SPOReplaceLink.sln")))
            dir = dir.Parent;
        return dir?.FullName ?? throw new InvalidOperationException("Repository root not found.");
    }

    [Fact]
    public void EnglishGuideHasTheSameSectionsAsTheUkrainianOne()
    {
        var uk = UserGuide.Headings(File.ReadAllText(Path.Combine(Root(), "docs", "USER_GUIDE.md")));
        var en = UserGuide.Headings(File.ReadAllText(Path.Combine(Root(), "docs", "USER_GUIDE.en.md")));
        Assert.Equal(uk.Select(h => h.Level), en.Select(h => h.Level));

        var source = UserGuide.For(L.Ukrainian);
        var english = UserGuide.For(L.English);
        Assert.Equal(L.English, english.Language);
        Assert.Equal(source.Pages.Select(p => p.Id), english.Pages.Select(p => p.Id));
        Assert.Equal(source.Parts.Count, english.Parts.Count);
        Assert.DoesNotContain(english.Pages, s => s.Title == "Contents");
        Assert.Equal("About this guide", english.Intro!.Title);
        var tab = english.Find("вкладка-продуктивність")!;
        Assert.Contains("performance.budgetPercent", tab.Html);
        Assert.DoesNotMatch(CyrillicRegex(), tab.Title);
    }

    [Fact]
    public void EveryLinkInsideTheEnglishGuideLeadsToASection()
    {
        string markdown = File.ReadAllText(Path.Combine(Root(), "docs", "USER_GUIDE.en.md"));
        var english = UserGuide.For(L.English);
        var anchors = AnchorRegex().Matches(markdown).Select(m => m.Groups[1].Value).Distinct().ToList();
        Assert.NotEmpty(anchors);
        // On GitHub the anchors are those of the English headings; in the app they lead to the shared section ids.
        var englishSlugs = UserGuide.Headings(markdown).Select(h => UserGuide.Slug(h.Title)).ToHashSet();
        Assert.All(anchors, a => Assert.Contains(a, englishSlugs));
        Assert.All(english.Pages, s => Assert.DoesNotContain("href=\"#", s.Html));
        var ids = english.Ids.ToHashSet();
        foreach (var page in english.Pages)
        {
            foreach (Match m in HelpLinkRegex().Matches(page.Html))
                Assert.True(ids.Contains(m.Groups[1].Value), $"{page.Id}: help/{m.Groups[1].Value}");
        }
    }

    [Fact]
    public void GuideFollowsTheLanguage()
    {
        using (L.Use(L.English))
            Assert.Equal(L.English, UserGuide.Instance.Language);
        using (L.Use(L.Ukrainian))
            Assert.Equal(L.Ukrainian, UserGuide.Instance.Language);
    }

    [GeneratedRegex(@"\]\(#([^)\s]+)\)")]
    private static partial Regex AnchorRegex();

    [GeneratedRegex("href=\"help/([^\"]+)\"")]
    private static partial Regex HelpLinkRegex();

    [GeneratedRegex(@"[Ѐ-ӿ]")]
    private static partial Regex CyrillicRegex();
}
