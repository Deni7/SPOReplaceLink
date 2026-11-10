using System.Diagnostics;
using Bunit;
using Microsoft.AspNetCore.Components;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging.Abstractions;
using MudBlazor;
using MudBlazor.Services;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Engine;
using SPOReplaceLink.Ui.Components.Pages;
using SPOReplaceLink.Ui.Components.Shared;
using SPOReplaceLink.Ui.Services;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;

namespace SPOReplaceLink.Host.Tests;

/// <summary>Components of the GUI rendered with bUnit over a real state store and the engine on the fake Graph.</summary>
public sealed class UiTests : BunitContext, IAsyncLifetime
{
    private const string Host = "https://" + FakeGraph.Host;
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "spo-rl-ui", Guid.NewGuid().ToString("N"));
    private readonly FakeGraph _graph = new();
    private StateStore _store = null!;
    private RunService _runs = null!;
    private SettingsService _settings = null!;

    /// <summary>Instead of the "Save as" dialog: writes to memory and remembers what was saved.</summary>
    private sealed class MemorySaver : IFileSaver
    {
        public Dictionary<string, byte[]> Saved { get; } = [];

        public async Task<string?> SaveAsync(string suggestedName, Func<Stream, Task> write)
        {
            using var ms = new MemoryStream();
            await write(ms);
            Saved[suggestedName] = ms.ToArray();
            return "memory:" + suggestedName;
        }
    }

    private MemorySaver Saver { get; } = new();

    public Task InitializeAsync()
    {
        _store = new StateStore(Path.Combine(_dir, "state.db"));
        var apps = new FakeAppFactory(_graph);
        _runs = new RunService(_store, apps, NullLoggerFactory.Instance, new EngineOptions
        {
            TempRoot = Path.Combine(_dir, "temp"),
            PollInterval = TimeSpan.FromMilliseconds(20),
        });
        var browser = new SpoBrowser(_runs, apps);
        _settings = new SettingsService(_runs, browser);
        Services.AddMudServices();
        Services.AddSingleton(_store);
        Services.AddSingleton(_runs);
        Services.AddSingleton(browser);
        Services.AddSingleton(_settings);
        Services.AddSingleton(new StatusService(_runs));
        Services.AddSingleton<IFileSaver>(Saver);
        JSInterop.Mode = JSRuntimeMode.Loose;
        return Task.CompletedTask;
    }

    async Task IAsyncLifetime.DisposeAsync()
    {
        await _runs.DisposeAsync();
        await base.DisposeAsync();
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try
        {
            Directory.Delete(_dir, true);
        }
        catch (IOException)
        {
        }
    }

    private IRenderedComponent<T> RenderWithProviders<T>(Action<ComponentParameterCollectionBuilder<T>>? parameters = null) where T : IComponent =>
        RenderPage(parameters).FindComponent<T>();

    /// <summary>The component with the providers around it; dialogs and snackbars render here, outside the component.</summary>
    private IRenderedComponent<Providers> RenderPage<T>(Action<ComponentParameterCollectionBuilder<T>>? parameters = null) where T : IComponent =>
        Render<Providers>(p => p.AddChildContent(parameters ?? (_ => { })));

    /// <summary>The MudBlazor providers followed by the child content.</summary>
    private sealed class Providers : ComponentBase
    {
        [Parameter]
        public RenderFragment? ChildContent { get; set; }

        protected override void BuildRenderTree(Microsoft.AspNetCore.Components.Rendering.RenderTreeBuilder b)
        {
            b.OpenComponent<MudPopoverProvider>(0);
            b.CloseComponent();
            b.OpenComponent<MudDialogProvider>(1);
            b.CloseComponent();
            b.OpenComponent<MudSnackbarProvider>(2);
            b.CloseComponent();
            b.AddContent(3, ChildContent);
        }
    }

    private void SeedSettings()
    {
        _store.PutSetting(SettingKinds.Connection, "app-a", new Connection
        {
            Name = "app-a", TenantId = "t", ClientId = Guid.NewGuid().ToString(), CertificateThumbprint = new string('A', 40),
        }.ToJson(), DateTimeOffset.UtcNow);
        _store.PutSetting(SettingKinds.Template, "finance", new Template
        {
            Name = "finance", Rules = [new RuleEntry { Old = Host + "/Document", New = Host + "/sites/finance" }],
        }.ToJson(), DateTimeOffset.UtcNow);
        var profile = new Profile
        {
            Name = "finance",
            BaseHost = Host,
            Templates = ["finance"],
            Targets = [new Target { SiteUrl = Host + "/sites/finance", Library = "Documents" }],
            Connections = ["app-a"],
        };
        profile.Retry.FilePausesMinutes = [0.0005];
        profile.Retry.RequestAttempts = 1;
        profile.Retry.FileAttempts = 1;
        _store.PutSetting(SettingKinds.Profile, "finance", profile.ToJson(), DateTimeOffset.UtcNow);
    }

    [Fact]
    public async Task DashboardCountersMatchTheDatabase()
    {
        SeedSettings();
        var lib = _graph.AddLibrary("/sites/finance", "Documents", "Shared Documents");
        for (int i = 0; i < 5; i++)
            _graph.AddFile(lib, $"Fin/a{i}.xlsx", TestXlsx.WithHyperlink(Host + "/Document/Fin/x.xlsx"));
        _graph.AddFile(lib, "Fin/ok.xlsx", TestXlsx.WithHyperlink(Host + "/sites/finance/ok.xlsx"));
        _graph.AddFile(lib, "Fin/broken.xlsx", "not a zip"u8.ToArray());
        long id = _runs.CreateRun(RunModes.Apply, "finance");
        Assert.Equal(RunStatuses.CompletedWithErrors, await _runs.Start(id).Completion.WaitAsync(TimeSpan.FromSeconds(30)));

        var cut = RenderWithProviders<RunDashboard>(p => p.Add(x => x.RunId, id));
        var counts = _store.CountByStatus(id);
        string text = cut.Markup;
        Assert.Contains("завершено з помилками", text);
        Assert.Contains($">{counts[FileStatuses.Done]}<", text.Replace(" ", ""));
        Assert.Contains("Оброблено <b>7</b> з <b>7</b>", text);
        Assert.Contains("Замін посилань", text);
        cut.WaitForAssertion(() => Assert.Contains("гіперпосилання: 5", cut.Markup), TimeSpan.FromSeconds(10));
        Assert.Contains("файл: 1", text); // errors by class
        // Targets: the library on its own line with the site under it, not a narrow table column.
        cut.WaitForAssertion(() => Assert.Contains("breakdown-row", cut.Markup), TimeSpan.FromSeconds(10));
        Assert.Contains($">{Host}/sites/finance<", cut.Markup);
        // Controls of a finished apply run: retry, verify and delete, no pause/stop.
        Assert.Contains("Перевірити (verify)", text);
        Assert.Contains("Видалити…", text);
        Assert.DoesNotContain("Пауза", text);
    }

    [Fact]
    public async Task DeleteButtonWarnsAboutWrittenFilesAndRemovesTheRun()
    {
        SeedSettings();
        var lib = _graph.AddLibrary("/sites/finance", "Documents", "Shared Documents");
        _graph.AddFile(lib, "Fin/a.xlsx", TestXlsx.WithHyperlink(Host + "/Document/Fin/x.xlsx"));
        long id = _runs.CreateRun(RunModes.Apply, "finance");
        await _runs.Start(id).Completion.WaitAsync(TimeSpan.FromSeconds(30));

        // The confirmation is shown by the dialog provider, outside the dashboard: look in the whole page.
        var root = RenderPage<RunDashboard>(p => p.Add(x => x.RunId, id));
        await root.InvokeAsync(() => root.FindAll("button").First(b => b.TextContent.Contains("Видалити…")).Click());
        root.WaitForAssertion(() => Assert.Contains("записав у SharePoint 1 файлів", root.Markup));
        Assert.NotNull(_store.GetRun(id));

        await root.InvokeAsync(() => root.FindAll("button").Last(b => b.TextContent.Trim() == "Видалити").Click());
        root.WaitForAssertion(() => Assert.Null(_store.GetRun(id)));
    }

    [Fact]
    public void RenameUpdatesProfilesThatUseIt()
    {
        SeedSettings();
        var updated = _settings.Rename(SettingKinds.Template, "finance", "finance-2024");
        Assert.Equal(["finance"], updated);
        Assert.Null(_settings.GetTemplate("finance"));
        Assert.Equal(["finance-2024"], _settings.GetProfile("finance")!.Templates);
        Assert.True(_runs.ValidateProfile(_settings.GetProfile("finance")!).IsValid);

        _settings.Rename(SettingKinds.Connection, "app-a", "app-main");
        Assert.Equal(["app-main"], _settings.GetProfile("finance")!.Connections);
        _settings.Rename(SettingKinds.Profile, "finance", "Fin");
        Assert.Equal("Fin", _settings.GetProfile("Fin")!.Name);
        _store.PutSetting(SettingKinds.Template, "other", new Template { Name = "other" }.ToJson(), DateTimeOffset.UtcNow);
        Assert.Throws<InvalidOperationException>(() => _settings.Rename(SettingKinds.Template, "finance-2024", "other"));
        Assert.Empty(_settings.Rename(SettingKinds.Profile, "Fin", "Fin"));
    }

    [Fact]
    public async Task SortsFilesOnTheServer()
    {
        SeedSettings();
        var config = new RunConfig { Profile = _runs.GetProfile("finance").ResolveTemplates(_runs.Templates()), Connections = [.. _runs.Connections()] };
        long id = SyntheticRun.Create(_store, Path.Combine(_dir, "state.db"), "finance", config.ToJson(), 500);
        var bySize = _store.QueryFiles(id, new FileFilter(), 0, 500, "size", descending: true);
        Assert.Equal(bySize.Select(f => f.Size).OrderByDescending(s => s), bySize.Select(f => f.Size));
        var byPath = _store.QueryFiles(id, new FileFilter(), 0, 50, "path");
        Assert.Equal(byPath.Select(f => f.Path).OrderBy(p => p, StringComparer.OrdinalIgnoreCase), byPath.Select(f => f.Path));
        // Unknown column: falls back to id, no SQL injection.
        Assert.Equal(_store.QueryFiles(id, new FileFilter(), 0, 5).Select(f => f.Id), _store.QueryFiles(id, new FileFilter(), 0, 5, "1; DROP TABLE files").Select(f => f.Id));
        await Task.CompletedTask;
    }

    [Fact]
    public async Task FilePageSavesTheVersionBeforeTheChange()
    {
        SeedSettings();
        var lib = _graph.AddLibrary("/sites/finance", "Documents", "Shared Documents");
        var original = TestXlsx.WithHyperlink(Host + "/Document/Fin/x.xlsx");
        var item = _graph.AddFile(lib, "Fin/a.xlsx", original);
        long id = _runs.CreateRun(RunModes.Apply, "finance");
        await _runs.Start(id).Completion.WaitAsync(TimeSpan.FromSeconds(30));
        var file = _store.QueryFiles(id, new FileFilter(), 0, 10).Single();
        Assert.Equal("1.0", file.VersionBefore);
        Assert.NotEqual(original, item.Content);

        var page = RenderWithProviders<FilePage>(p => p.Add(x => x.Id, file.Id));
        await page.InvokeAsync(() => page.FindAll("button").First(b => b.TextContent.Contains("Зберегти версію до зміни")).Click());
        page.WaitForAssertion(() => Assert.True(Saver.Saved.ContainsKey("a (версія 1.0).xlsx")), TimeSpan.FromSeconds(10));
        Assert.Equal(original, Saver.Saved["a (версія 1.0).xlsx"]);
    }

    [Fact]
    public async Task NewRunDialogAppliesOneOffChangesWithoutSavingThem()
    {
        SeedSettings();
        _graph.AddLibrary("/sites/finance", "Documents", "Shared Documents");
        var root = Render(b =>
        {
            b.OpenComponent<MudPopoverProvider>(0);
            b.CloseComponent();
            b.OpenComponent<MudDialogProvider>(1);
            b.CloseComponent();
        });
        var dialogs = Services.GetRequiredService<IDialogService>();
        IDialogReference? reference = null;
        await root.InvokeAsync(async () => reference = await dialogs.ShowAsync<NewRunDialog>("Новий прогін", new DialogParameters<NewRunDialog> { { x => x.Profile, "finance" } }));
        root.WaitForAssertion(() => Assert.Contains("Разові зміни", root.Markup));

        // The first numeric field of the one-off panel is "Завантажень".
        var downloads = root.FindAll("input").First(i => i.GetAttribute("value") == "4" || i.GetAttribute("value") == "8");
        downloads.Change("2");
        root.WaitForAssertion(() => Assert.Contains("(є зміни)", root.Markup));
        await root.InvokeAsync(() => root.FindAll("button").First(b => b.TextContent.Trim() == "Запустити").Click());

        var result = await reference!.Result;
        long runId = Assert.IsType<long>(result!.Data);
        var snapshot = RunConfig.Parse(_store.GetRun(runId)!.ConfigSnapshot);
        Assert.Equal(2, snapshot.Profile.Performance.Downloads);
        Assert.NotEqual(2, _settings.GetProfile("finance")!.Performance.Downloads);
    }

    [Fact]
    public void ProfileEditorMarksErrorsOnTheirTabsWhenOpened()
    {
        SeedSettings();
        var p = Profile.Parse(_store.GetSetting(SettingKinds.Profile, "finance")!);
        p.Connections = ["missing"];
        p.Filters.MinFileSizeKb = -1;
        _store.PutSetting(SettingKinds.Profile, "finance", p.ToJson(), DateTimeOffset.UtcNow);
        var cut = RenderWithProviders<ProfileEdit>(x => x.Add(y => y.Name, "finance"));
        Assert.Contains("Помилок: 2", cut.Markup);
        // Each error is next to its section (the badge with the count is on the tab).
        Assert.Equal(2, cut.FindAll(".mud-tab .mud-badge").Count);
    }

    [Fact]
    public async Task ExportsAreSavedFromThePages()
    {
        SeedSettings();
        var lib = _graph.AddLibrary("/sites/finance", "Documents", "Shared Documents");
        _graph.AddFile(lib, "Fin/a.xlsx", TestXlsx.WithHyperlink(Host + "/Document/Fin/x.xlsx"));
        long id = _runs.CreateRun(RunModes.Scan, "finance");
        await _runs.Start(id).Completion.WaitAsync(TimeSpan.FromSeconds(30));

        var table = RenderWithProviders<FilesTable>(p => p.Add(x => x.RunId, id));
        await table.InvokeAsync(() => table.FindAll("button").First(b => b.TextContent.Trim() == "XLSX").Click());
        table.WaitForAssertion(() => Assert.True(Saver.Saved.ContainsKey($"run-{id}-files.xlsx")));
        var rows = OpenXml.SimpleXlsxReader.ReadFirstSheet(new MemoryStream(Saver.Saved[$"run-{id}-files.xlsx"]));
        Assert.Equal("Fin/a.xlsx", rows[1][1]);

        await DisposeComponentsAsync();
        var profile = RenderWithProviders<ProfileEdit>(p => p.Add(x => x.Name, "finance"));
        await profile.InvokeAsync(() => profile.FindAll("button").First(b => b.TextContent.Trim() == "JSON").Click());
        profile.WaitForAssertion(() => Assert.True(Saver.Saved.ContainsKey("finance.profile.json")));
        Assert.Contains("\"targets\"", System.Text.Encoding.UTF8.GetString(Saver.Saved["finance.profile.json"]));
    }

    [Fact]
    public void RulesEditorShowsConflictsWhileTyping()
    {
        var rules = new List<RuleEntry> { new() { Old = Host + "/Document", New = Host + "/sites/b" } };
        var preceding = new List<RuleEntry> { new() { Old = Host + "/Document/", New = Host + "/sites/a" } };
        var cut = RenderWithProviders<RulesEditor>(p => p.Add(x => x.Rules, rules).Add(x => x.Preceding, preceding).Add(x => x.BaseHost, Host));
        Assert.Contains("Конфлікт з правилом 1", cut.Markup);

        rules[0].Old = Host + "/Other";
        cut.Render(p => p.Add(x => x.Rules, rules));
        Assert.DoesNotContain("Конфлікт", cut.Markup);
        Assert.Contains("Набір правил коректний", cut.Markup);
    }

    [Fact]
    public void RulesEditorTestsALink()
    {
        var rules = new List<RuleEntry> { new() { Old = "/Document/", New = Host + "/sites/finance/" } };
        var cut = RenderWithProviders<RulesEditor>(p => p.Add(x => x.Rules, rules).Add(x => x.BaseHost, Host));
        var inputs = cut.FindAll("input");
        // inputs: old, new, name of the rule; then test url and test folder
        inputs[3].Input(Host + "/Document/Fin/Звіт 1.xlsx");
        cut.FindAll("button").First(b => b.TextContent.Contains("Перевірити")).Click();
        Assert.Contains(Host + "/sites/finance/Fin/Звіт 1.xlsx", cut.Markup);
    }

    [Fact]
    public async Task ProfileEditorSavesAndShowsValidation()
    {
        SeedSettings();
        var cut = RenderWithProviders<ProfileEdit>(p => p.Add(x => x.Name, "finance"));
        Assert.Contains("Профіль «finance»", cut.Markup);
        cut.FindAll("button").First(b => b.TextContent.Contains("Зберегти")).Click();
        Assert.Contains("Помилок немає", cut.Markup);
        Assert.NotNull(_store.GetSetting(SettingKinds.Profile, "finance"));

        // An unknown connection makes the profile invalid; the editor says why.
        var p = Profile.Parse(_store.GetSetting(SettingKinds.Profile, "finance")!);
        p.Connections = ["missing"];
        _store.PutSetting(SettingKinds.Profile, "finance", p.ToJson(), DateTimeOffset.UtcNow);
        await DisposeComponentsAsync();
        cut = RenderWithProviders<ProfileEdit>(x => x.Add(y => y.Name, "finance"));
        cut.FindAll("button").First(b => b.TextContent.Contains("Перевірити")).Click();
        Assert.Contains("Помилок: 1", cut.Markup);
        Assert.Single(cut.FindAll(".mud-tab .mud-badge"));
    }

    [Fact]
    public async Task ErrorsViewGroupsFailures()
    {
        SeedSettings();
        var config = new RunConfig { Profile = _runs.GetProfile("finance").ResolveTemplates(_runs.Templates()), Connections = [.. _runs.Connections()] };
        long id = SyntheticRun.Create(_store, Path.Combine(_dir, "state.db"), "finance", config.ToJson(), 2000);
        var cut = RenderWithProviders<ErrorsView>(p => p.Add(x => x.RunId, id));
        Assert.Contains("failed:transient", cut.Markup);
        Assert.Contains("failed:locked", cut.Markup);
        Assert.Contains("retry пізніше", cut.Markup);

        // The "Помилки" card opens only failed files; "Потребують уваги" — what that card counts.
        cut.Render(p => p.Add(x => x.Scope, StateStore.ErrorScope.Failed));
        Assert.Contains("failed:file", cut.Markup);
        Assert.DoesNotContain("needs_excel", cut.Markup);
        Assert.DoesNotContain("done:residual", cut.Markup);
        cut.Render(p => p.Add(x => x.Scope, StateStore.ErrorScope.Attention));
        Assert.Contains("needs_excel", cut.Markup);
        Assert.Contains("done:residual", cut.Markup);
        Assert.Equal(_store.CountAttention(id), _store.GroupErrors(id, StateStore.ErrorScope.Attention).Sum(g => g.Count));

        // The "attention" filter value takes exactly what the card counts, alone or with other statuses, and so does a retry.
        long attention = _store.CountAttention(id);
        Assert.True(attention > 0);
        Assert.Equal(attention, _store.CountFiles(id, new FileFilter { Statuses = [FileStatuses.Attention] }));
        Assert.True(_store.CountFiles(id, new FileFilter { Statuses = [FileStatuses.Attention, FileStatuses.Unchanged] }) > attention);

        await DisposeComponentsAsync();
        var dashboard = RenderWithProviders<RunDashboard>(p => p.Add(x => x.RunId, id));
        Assert.Contains($"href=\"runs/{id}?tab=errors&amp;show=failed\"", dashboard.Markup);
        Assert.Contains($"href=\"runs/{id}?tab=errors&amp;show=attention\"", dashboard.Markup);

        // The tab and the filter go to the address, so the same card link works again after switching tabs by hand.
        await DisposeComponentsAsync();
        var nav = Services.GetRequiredService<Microsoft.AspNetCore.Components.NavigationManager>();
        var page = RenderWithProviders<RunPage>(p => p.Add(x => x.Id, id));
        page.WaitForAssertion(() => Assert.Contains(page.FindAll(".mud-tab"), t => t.TextContent.Contains("Помилки й увага")));
        page.FindAll(".mud-tab").First(t => t.TextContent.Contains("Помилки й увага")).Click();
        page.WaitForAssertion(() => Assert.EndsWith("?tab=errors", nav.Uri));
        page.WaitForAssertion(() => Assert.Contains(page.FindAll("button"), b => b.TextContent.Trim() == "Помилки"));
        page.FindAll("button").First(b => b.TextContent.Trim() == "Помилки").Click();
        page.WaitForAssertion(() => Assert.EndsWith("?tab=errors&show=failed", nav.Uri));
        page.FindAll(".mud-tab").First(t => t.TextContent.Contains("Огляд")).Click();
        page.WaitForAssertion(() => Assert.DoesNotContain("tab=", nav.Uri));

        long retry = _store.CreateRun(RunModes.Apply, id, "finance", config.ToJson(), enumerate: false, targets: 0, DateTimeOffset.UtcNow);
        Assert.Equal(attention, _store.CopyFiles(id, retry, new FileFilter { Statuses = [FileStatuses.Attention] }, DateTimeOffset.UtcNow));
    }

    /// <summary>Stage 3, item 12: the GUI stays responsive on a run of 200 000 files (criterion: under 1 s).</summary>
    [Fact]
    [Trait("Category", "Performance")]
    public void LargeRunStaysResponsive()
    {
        SeedSettings();
        var config = new RunConfig { Profile = _runs.GetProfile("finance").ResolveTemplates(_runs.Templates()), Connections = [.. _runs.Connections()] };
        long id = SyntheticRun.Create(_store, Path.Combine(_dir, "state.db"), "finance", config.ToJson(), 200_000);

        void Within(string what, Action action)
        {
            action(); // warm-up: first query compiles statements and fills the page cache
            var sw = Stopwatch.StartNew();
            action();
            Assert.True(sw.ElapsedMilliseconds < 1000, $"{what}: {sw.ElapsedMilliseconds} ms");
        }

        var status = new StatusService(_runs);
        Within("стан прогону", () => status.Get(id, fresh: true));
        Within("сторінка файлів", () => _store.QueryFiles(id, new FileFilter { Statuses = ["failed"] }, 5000, 50));
        Within("пошук", () => _store.QueryFiles(id, new FileFilter { Search = "0012" }, 0, 50));
        Within("групи помилок", () => _store.GroupErrors(id));
        Within("розбивки", () => { _store.BreakdownBy(id, "ext"); _store.BreakdownBy(id, "target_idx"); });
        Within("заміни за місцями", () => _store.ReplacementsByKind(id));
        Within("дашборд", () =>
        {
            DisposeComponentsAsync().GetAwaiter().GetResult();
            RenderWithProviders<RunDashboard>(p => p.Add(x => x.RunId, id));
        });
    }
}
