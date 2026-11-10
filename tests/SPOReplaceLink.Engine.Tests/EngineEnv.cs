using Microsoft.Extensions.Logging.Abstractions;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;

namespace SPOReplaceLink.Engine.Tests;

/// <summary>A fake tenant, a state store in a temp folder and a run service over them.</summary>
internal sealed class EngineEnv : IAsyncDisposable
{
    public const string Host = "https://" + FakeGraph.Host;
    public const string OldUrl = Host + "/Document/Fin/a.xlsx";
    public const string NewUrl = Host + "/sites/finance/Fin/a.xlsx";

    /// <param name="excel">Excel pool factory; null — runs without Excel.</param>
    public EngineEnv(Func<string, Func<int, TimeSpan, ExcelWorker.IExcelConverter>>? excel = null)
    {
        Dir = Path.Combine(Path.GetTempPath(), "spo-rl-engine", Guid.NewGuid().ToString("N"));
        Store = new StateStore(Path.Combine(Dir, "state.db"));
        Apps = new FakeAppFactory(Graph);
        Service = new RunService(Store, Apps, NullLoggerFactory.Instance, new EngineOptions
        {
            TempRoot = Path.Combine(Dir, "temp"),
            PollInterval = TimeSpan.FromMilliseconds(20),
            EnumerationRetryDelay = TimeSpan.FromMilliseconds(50),
            Excel = excel?.Invoke(Path.Combine(Dir, "temp")),
        });
        Library = Graph.AddLibrary("/sites/finance", "Documents", "Shared Documents");
        AddConnection("app-a");
    }

    public string Dir { get; }
    public FakeGraph Graph { get; } = new();
    public FakeAppFactory Apps { get; }
    public StateStore Store { get; }
    public RunService Service { get; }
    public FakeGraph.FakeDrive Library { get; }

    public void AddConnection(string name, LicenseTier tier = LicenseTier.UpTo1000, int? tenantOverride = null)
    {
        var c = new Connection
        {
            Name = name,
            TenantId = "contoso.onmicrosoft.com",
            ClientId = Guid.NewGuid().ToString(),
            CertificateThumbprint = new string('A', 40),
            Licenses = tier,
            TenantPer5MinOverride = tenantOverride,
        };
        Store.PutSetting(SettingKinds.Connection, name, c.ToJson(), DateTimeOffset.UtcNow);
    }

    public static Profile NewProfile(params string[] connections)
    {
        var p = new Profile
        {
            Name = "test",
            BaseHost = Host,
            Rules = [new RuleEntry { Old = Host + "/Document", New = Host + "/sites/finance" }],
            Targets = [new Target { SiteUrl = Host + "/sites/finance", Library = "Documents" }],
            Connections = connections.Length == 0 ? ["app-a"] : [.. connections],
        };
        p.Retry.FilePausesMinutes = [0.0005];
        p.Retry.LockedWaveMinutes = 0.001;
        p.Retry.LockedLimitMinutes = 1;
        p.Retry.BreakerWindow = 1000;
        p.Performance.Downloads = 4;
        p.Performance.Uploads = 2;
        p.Performance.Processors = 2;
        return p;
    }

    public void SaveProfile(Profile p) => Store.PutSetting(SettingKinds.Profile, p.Name!, p.ToJson(), DateTimeOffset.UtcNow);

    public long Create(string mode, Profile? profile = null)
    {
        var p = profile ?? NewProfile();
        SaveProfile(p);
        return Service.CreateRun(mode, p.Name);
    }

    public async Task<string> RunAsync(long runId, int timeoutSeconds = 60) =>
        await Service.Start(runId).Completion.WaitAsync(TimeSpan.FromSeconds(timeoutSeconds));

    public FakeGraph.FakeItem AddXlsx(string path, string url) => Graph.AddFile(Library, path, TestXlsx.WithHyperlink(url));

    public IReadOnlyList<FileRow> Files(long runId) => Store.QueryFiles(runId, new FileFilter(), 0, 10_000);

    public string StatusOf(long runId, string path) => Files(runId).Single(f => f.Path == path).StatusCode!;

    public async ValueTask DisposeAsync()
    {
        await Service.DisposeAsync();
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try
        {
            Directory.Delete(Dir, true);
        }
        catch (IOException)
        {
        }
    }
}
