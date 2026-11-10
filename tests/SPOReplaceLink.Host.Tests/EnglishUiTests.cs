using Bunit;
using MudBlazor.Services;
using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.Ui.Components.Pages;

namespace SPOReplaceLink.Host.Tests;

/// <summary>Pages render in English when the language is English: the guide, labels and messages of the page.</summary>
public sealed class EnglishUiTests : BunitContext
{
    [Fact]
    public void HelpPageRendersInEnglish()
    {
        using var _ = L.Use(L.English);
        Services.AddMudServices();
        JSInterop.Mode = JSRuntimeMode.Loose;

        var page = Render<HelpPage>(p => p.Add(x => x.Id, "вкладка-фільтри"));
        Assert.Contains("filters.maxFileSizeMb", page.Markup);
        Assert.Contains("help/вкладка-цілі", page.Markup); // same section ids in both languages
        Assert.DoesNotContain("Вкладка «Фільтри»", page.Markup);

        var missing = Render<HelpPage>(p => p.Add(x => x.Id, "nope"));
        Assert.DoesNotContain("Розділ не знайдено", missing.Markup);
    }
}
