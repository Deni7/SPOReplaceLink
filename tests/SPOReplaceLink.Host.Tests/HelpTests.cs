using System.Text.RegularExpressions;
using Bunit;
using MudBlazor.Services;
using SPOReplaceLink.Ui.Components.Pages;
using SPOReplaceLink.Ui.Services;

namespace SPOReplaceLink.Host.Tests;

/// <summary>The user guide as in-app help: sections, links between them and the "?" buttons of the screens.</summary>
public sealed partial class HelpTests : BunitContext
{
    private static readonly UserGuide Guide = UserGuide.Instance;

    private static string Root()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "SPOReplaceLink.sln")))
            dir = dir.Parent;
        return dir?.FullName ?? throw new InvalidOperationException("Корінь репозиторію не знайдено.");
    }

    [Fact]
    public void GuideIsSplitIntoPartsAndSections()
    {
        Assert.Equal(3, Guide.Parts.Count);
        Assert.NotNull(Guide.Intro);
        Assert.DoesNotContain(Guide.Pages, s => s.Title == "Зміст");
        var profile = Guide.Find("33-профіль")!;
        Assert.Equal("Частина 3. Довідник", profile.Group);
        Assert.Contains(profile.Children, c => c.Id == "вкладка-продуктивність");
        var tab = Guide.Find("вкладка-продуктивність")!;
        Assert.Equal("33-профіль", tab.ParentId);
        Assert.Contains("<table>", tab.Html);
        Assert.Contains("performance.budgetPercent", tab.Html);
        // Code blocks with "#" inside are not headings.
        Assert.DoesNotContain(Guide.Pages, s => s.Title.StartsWith("include", StringComparison.Ordinal));
        Assert.True(Guide.Pages.Count > 30, $"{Guide.Pages.Count} сторінок");
    }

    [Fact]
    public void EveryLinkInsideTheGuideLeadsToASection()
    {
        string markdown = File.ReadAllText(Path.Combine(Root(), "docs", "USER_GUIDE.md"));
        var anchors = AnchorRegex().Matches(markdown).Select(m => m.Groups[1].Value).Distinct().ToList();
        Assert.NotEmpty(anchors);
        Assert.All(anchors, a => Assert.True(Guide.Find(a) is not null, $"посилання #{a} не веде на розділ"));
        // In the app they are links to help pages; links to other repository files keep only their text.
        Assert.All(Guide.Pages, s => Assert.DoesNotContain("href=\"#", s.Html));
        Assert.All(Guide.Pages, s => Assert.DoesNotContain("href=\"../", s.Html));
    }

    [Fact]
    public void EveryHelpButtonOfTheScreensLeadsToASection()
    {
        var ids = Directory.EnumerateFiles(Path.Combine(Root(), "src", "SPOReplaceLink.Ui"), "*.razor", SearchOption.AllDirectories)
            .SelectMany(f => ButtonRegex().Matches(File.ReadAllText(f)).Select(m => (File: Path.GetFileName(f), Id: m.Groups.Values.Skip(1).First(g => g.Success).Value)))
            .Where(x => !x.Id.StartsWith('@'))
            .ToList();
        Assert.True(ids.Count >= 15, $"кнопок довідки: {ids.Count}");
        Assert.All(ids, x => Assert.True(Guide.Find(x.Id) is not null, $"{x.File}: розділ «{x.Id}» не знайдено"));
    }

    [Fact]
    public void SearchFindsSectionsByWords()
    {
        var found = Guide.Search("бюджет RU");
        Assert.NotEmpty(found);
        Assert.Contains(found, s => s.Id == "вкладка-продуктивність");
        Assert.Empty(Guide.Search("неіснуюче-слово-xyz"));
    }

    [Fact]
    public void HelpPageShowsTheSectionAndItsNeighbours()
    {
        Services.AddMudServices();
        JSInterop.Mode = JSRuntimeMode.Loose;
        var page = Render<HelpPage>(p => p.Add(x => x.Id, "вкладка-фільтри"));
        Assert.Contains("Вкладка «Фільтри»", page.Markup);
        Assert.Contains("filters.maxFileSizeMb", page.Markup);
        Assert.Contains("help/вкладка-цілі", page.Markup); // previous
        Assert.Contains("help/вкладка-місця-заміни", page.Markup); // next

        var missing = Render<HelpPage>(p => p.Add(x => x.Id, "nope"));
        Assert.Contains("Розділ не знайдено", missing.Markup);
    }

    [GeneratedRegex(@"\]\(#([^)\s]+)\)")]
    private static partial Regex AnchorRegex();

    [GeneratedRegex("<HelpButton[^>]*Section=\"([^\"]+)\"|\\sHelp=\"([^\"]+)\"|href=\"help/([^\"@]+)\"")]
    private static partial Regex ButtonRegex();
}
