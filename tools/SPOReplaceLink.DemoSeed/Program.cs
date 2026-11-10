// Fills a data folder with demo settings and runs, made by the real engine against the fake Graph,
// so the app can be looked at without SharePoint.
// Usage: DemoSeed <dataDir> [--large N] [--open] [--port PORT]
// --open starts the desktop app on these data with the same fake Graph, so new runs can be started from its window;
// --port is the local API port for the CLI (default 5097).
// The fake SharePoint is saved to <dataDir>\fake-sharepoint.json when the app closes, so the same folder can be
// opened again later (e.g. to continue a stopped run).
using System.Net;
using Microsoft.Extensions.Logging.Abstractions;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Engine;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;

string dir = Path.GetFullPath(args.Length > 0 ? args[0] : "out/demo-data");
int large = Option("--large") is { } l ? int.Parse(l) : 0;
bool open = args.Contains("--open") || args.Contains("--serve");
int port = (Option("--serve") ?? Option("--port")) is { } s && int.TryParse(s, out int parsed) ? parsed : 5097;
string dbPath = Path.Combine(dir, "state.db");
string graphPath = Path.Combine(dir, "fake-sharepoint.json");
Directory.CreateDirectory(dir);

FakeGraph graph;
if (File.Exists(dbPath))
{
    if (!File.Exists(graphPath))
    {
        Console.Error.WriteLine($"Папка {dir} містить демо-дані без збереженого емульованого SharePoint (створена старою версією утиліти). Вкажіть нову папку.");
        return 2;
    }
    if (!open)
    {
        Console.Error.WriteLine($"Папка {dir} уже містить демо-дані. Щоб відкрити їх, додайте --open; щоб створити нові — вкажіть іншу папку.");
        return 2;
    }
    graph = FakeGraph.Load(graphPath);
    Console.WriteLine($"Відкриваю наявні демо-дані з {dir}…");
}
else
{
    graph = await Seed();
    graph.Save(graphPath);
}

if (!open)
    return 0;

graph.Faults.Clear();
Console.WriteLine("Відкриваю застосунок SPOReplaceLink на цих даних (SharePoint емулюється)…");
// WPF needs an STA thread; the top-level program is not one.
int exit = 0;
var ui = new Thread(() => exit = SPOReplaceLink.Desktop.DesktopApp.Run(
    ["--data", dir, "--urls", $"http://localhost:{port}", "--continueInterrupted", "false", .. Option("--selftest") is { } st ? new[] { "--selftest", st } : []],
    services =>
    {
        services.AddSingleton<SPOReplaceLink.Spo.ISpoAppFactory>(new FakeAppFactory(graph));
        services.AddSingleton(new EngineOptions
        {
            TempRoot = Path.Combine(dir, "temp"),
            PollInterval = TimeSpan.FromMilliseconds(50),
            MetricsInterval = TimeSpan.FromSeconds(5),
            Excel = SPOReplaceLink.ExcelWorker.ExcelPool.Factory(Path.Combine(dir, "temp")),
        });
    }));
ui.SetApartmentState(ApartmentState.STA);
ui.Start();
ui.Join();
// What the runs wrote to the fake SharePoint is kept for the next opening of this folder.
graph.Save(graphPath);
Console.WriteLine($"Емульований SharePoint збережено: {graphPath}");
return exit;

async Task<FakeGraph> Seed()
{
    Console.WriteLine($"Готую демо-дані в {dir}{(large > 0 ? $" (разом із синтетичним прогоном на {large} файлів)" : "")}…");
    const string Host = "https://" + FakeGraph.Host;
    var store = new StateStore(dbPath);
    var g = new FakeGraph { DeltaPageSize = 50 };
    var library = g.AddLibrary("/sites/finance", "Documents", "Shared Documents");
    var archive = g.AddLibrary("/sites/finance", "Архів", "Archive");
    var now = DateTimeOffset.UtcNow;

    store.PutSetting(SettingKinds.Connection, "demo-app", new Connection
    {
        Name = "demo-app",
        TenantId = "contoso.onmicrosoft.com",
        ClientId = "11111111-2222-3333-4444-555555555555",
        CertificateThumbprint = new string('A', 40),
        Licenses = LicenseTier.UpTo5000,
    }.ToJson(), now);
    store.PutSetting(SettingKinds.Template, "finance", new Template
    {
        Name = "finance",
        Description = "Бібліотека Document перенесена на сайт finance",
        Rules =
        [
            new RuleEntry { Old = "/Document/", New = Host + "/sites/finance/", Name = "Document → finance" },
            new RuleEntry { Old = "/Document/Analytics/", New = "/sites/finance/analytics_migrated/", Name = "Analytics" },
        ],
    }.ToJson(), now);
    var profile = new Profile
    {
        Name = "finance",
        BaseHost = Host,
        Templates = ["finance"],
        Targets =
        [
            new Target { SiteUrl = Host + "/sites/finance", Library = "Documents", OldLocation = Host + "/Document" },
            new Target { SiteUrl = Host + "/sites/finance", Library = "Архів" },
        ],
        Connections = ["demo-app"],
    };
    profile.Retry.FilePausesMinutes = [0.001];
    profile.Retry.LockedWaveMinutes = 0.002;
    profile.Retry.LockedLimitMinutes = 0.01;
    profile.Retry.RequestAttempts = 2;
    store.PutSetting(SettingKinds.Profile, "finance", profile.ToJson(), now);

    for (int i = 0; i < 120; i++)
    {
        string folder = i % 3 == 0 ? "Fin/2024" : i % 3 == 1 ? "Analytics" : "Бухгалтерія";
        string url = i % 4 == 0 ? Host + "/sites/finance/ok.xlsx" : i % 5 == 0 ? "../Analytics/src.xlsx" : $"{Host}/Document/{folder}/Звіт%20{i}.xlsx";
        string text = i % 17 == 0 ? $"Див. {Host}/Document/Fin/old.xlsx" : "text";
        var item = g.AddFile(i % 6 == 0 ? archive : library, $"{folder}/Звіт {i:000}.xlsx", TestXlsx.WithHyperlink(url, text));
        if (i % 23 == 0)
            item.CheckedOutBy = "ivan@contoso.com";
    }
    g.AddFile(library, "Fin/broken.xlsx", "not a zip"u8.ToArray());
    g.AddFile(library, "Fin/old.xls", [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0]);
    // Real xls made by Excel (stage 4): links in a hyperlink and a HYPERLINK formula, fixed by ExcelWorker.
    if (await RealXlsAsync($"{Host}/Document/Fin/Звіт xls.xlsx") is { } xls)
        g.AddFile(library, "Fin/Звіт старий.xls", xls);
    g.Faults.Add(r => r.ItemName is { } n && n.EndsWith("013.xlsx", StringComparison.Ordinal) && r.Path.Contains("/content", StringComparison.Ordinal)
        ? FakeGraph.Error(HttpStatusCode.ServiceUnavailable, "serverBusy", 0)
        : null);

    var service = new RunService(store, new FakeAppFactory(g), NullLoggerFactory.Instance, new EngineOptions
    {
        TempRoot = Path.Combine(dir, "temp"),
        Excel = SPOReplaceLink.ExcelWorker.ExcelPool.Factory(Path.Combine(dir, "temp")),
        PollInterval = TimeSpan.FromMilliseconds(20),
        MetricsInterval = TimeSpan.FromMilliseconds(300),
    });
    foreach (string mode in new[] { RunModes.Inventory, RunModes.Scan, RunModes.Apply })
    {
        long id = service.CreateRun(mode, "finance");
        string status = await service.Start(id).Completion;
        Console.WriteLine($"{mode}: прогін {id} — {status}");
    }
    long verify = service.CreateChildRun(3, new FileFilter { Statuses = [FileStatuses.Done] }, RunModes.Verify);
    Console.WriteLine($"verify: прогін {verify} — {await service.Start(verify).Completion}");

    if (large > 0)
    {
        var config = new RunConfig { Profile = profile.ResolveTemplates(service.Templates()), Connections = [.. service.Connections()] };
        var sw = System.Diagnostics.Stopwatch.StartNew();
        long id = SyntheticRun.Create(store, dbPath, "finance-large", config.ToJson(), large);
        Console.WriteLine($"синтетичний прогін {id}: {large} файлів за {sw.Elapsed.TotalSeconds:0} с");
    }
    await service.DisposeAsync();
    Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
    Console.WriteLine("Готово: " + dir);
    return g;
}

/// <summary>An xls with a link, saved by Excel; null without Excel.</summary>
async Task<byte[]?> RealXlsAsync(string url)
{
    if (!OperatingSystem.IsWindows() || Type.GetTypeFromProgID("Excel.Application") is null)
        return null;
    string work = Path.Combine(dir, "temp", "demo-xls");
    Directory.CreateDirectory(work);
    string xlsx = Path.Combine(work, "src.xlsx");
    string xls = Path.Combine(work, "src.xls");
    const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
    using (var zip = System.IO.Compression.ZipFile.Open(xlsx, System.IO.Compression.ZipArchiveMode.Create))
    {
        void Add(string name, string text)
        {
            using var w = new StreamWriter(zip.CreateEntry(name).Open());
            w.Write(text);
        }
        Add("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
        Add("_rels/.rels", $"<Relationships xmlns=\"{PkgRel}\"><Relationship Id=\"rId1\" Type=\"{Rel}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
        Add("xl/workbook.xml", $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
        Add("xl/_rels/workbook.xml.rels", $"<Relationships xmlns=\"{PkgRel}\"><Relationship Id=\"rId1\" Type=\"{Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
        Add("xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheetData>" +
            "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Звіт</t></is></c></row>" +
            $"<row r=\"2\"><c r=\"A2\" t=\"str\"><f>HYPERLINK(\"{url}\",\"формула\")</f><v>формула</v></c></row>" +
            "</sheetData><hyperlinks><hyperlink ref=\"A1\" r:id=\"rId1\"/></hyperlinks></worksheet>");
        Add("xl/worksheets/_rels/sheet1.xml.rels", $"<Relationships xmlns=\"{PkgRel}\"><Relationship Id=\"rId1\" Type=\"{Rel}/hyperlink\" Target=\"{url.Replace(" ", "%20")}\" TargetMode=\"External\"/></Relationships>");
    }
    try
    {
        await using var pool = new SPOReplaceLink.ExcelWorker.ExcelPool(new(1, TimeSpan.FromMinutes(2), Path.Combine(dir, "temp")));
        await pool.ConvertAsync(xlsx, xls, SPOReplaceLink.ExcelWorker.XlFormat.Xls, CancellationToken.None);
        return await File.ReadAllBytesAsync(xls);
    }
    catch (SPOReplaceLink.ExcelWorker.ExcelException ex)
    {
        Console.WriteLine("Excel недоступний, демо без справжнього xls: " + ex.Message);
        return null;
    }
    finally
    {
        Directory.Delete(work, true);
    }
}

string? Option(string name)
{
    int i = Array.IndexOf(args, name);
    return i >= 0 && i + 1 < args.Length ? args[i + 1] : null;
}
