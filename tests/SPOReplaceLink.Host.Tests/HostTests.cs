using System.Net;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.DependencyInjection;
using SPOReplaceLink.Cli;
using SPOReplaceLink.Engine;
using SPOReplaceLink.Spo;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;

namespace SPOReplaceLink.Host.Tests;

/// <summary>CLI → host REST API → engine → fake SharePoint: the way an administrator runs the tool.</summary>
public sealed class HostTests : IAsyncLifetime
{
    private const string Host = "https://" + FakeGraph.Host;
    private const string OldUrl = Host + "/Document/Fin/a.xlsx";
    private const string NewUrl = Host + "/sites/finance/Fin/a.xlsx";

    private readonly string _dir = Path.Combine(Path.GetTempPath(), "spo-rl-host", Guid.NewGuid().ToString("N"));
    private readonly FakeGraph _graph = new();
    private FakeGraph.FakeDrive _library = null!;
    private WebApplicationFactory<HostEntry> _factory = null!;
    private readonly StringWriter _out = new();

    public Task InitializeAsync()
    {
        Directory.CreateDirectory(_dir);
        _library = _graph.AddLibrary("/sites/finance", "Documents", "Shared Documents");
        _factory = CreateFactory(continueInterrupted: false);
        return Task.CompletedTask;
    }

    private WebApplicationFactory<HostEntry> CreateFactory(bool continueInterrupted) => new WebApplicationFactory<HostEntry>().WithWebHostBuilder(b =>
    {
        b.UseSetting("data", _dir);
        b.UseSetting("continueInterrupted", continueInterrupted ? "true" : "false");
        b.ConfigureTestServices(s =>
        {
            s.AddSingleton<ISpoAppFactory>(new FakeAppFactory(_graph));
            s.AddSingleton(new EngineOptions { TempRoot = Path.Combine(_dir, "temp"), PollInterval = TimeSpan.FromMilliseconds(20) });
        });
    });

    public async Task DisposeAsync()
    {
        await _factory.DisposeAsync();
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try
        {
            Directory.Delete(_dir, true);
        }
        catch (IOException)
        {
        }
    }

    private HostCommands Cli(WebApplicationFactory<HostEntry>? factory = null) =>
        new(new HostClient(Http(factory)), _out);

    /// <summary>A client with the key of the instance, as the CLI of the same user has it.</summary>
    private HttpClient Http(WebApplicationFactory<HostEntry>? factory = null)
    {
        var f = factory ?? _factory;
        var http = f.CreateClient();
        http.DefaultRequestHeaders.Add(Core.Hosting.ApiEndpoint.KeyHeader, f.Services.GetRequiredService<ApiAccess>().Key);
        return http;
    }

    [Fact]
    public async Task ApiRefusesRequestsWithoutTheKeyOfThisInstance()
    {
        var anonymous = _factory.CreateClient();
        Assert.Equal(HttpStatusCode.Unauthorized, (await anonymous.GetAsync("api/runs")).StatusCode);
        anonymous.DefaultRequestHeaders.Add(Core.Hosting.ApiEndpoint.KeyHeader, new string('0', 64));
        Assert.Equal(HttpStatusCode.Unauthorized, (await anonymous.GetAsync("api/runs")).StatusCode);
        Assert.True((await Http().GetAsync("api/runs")).IsSuccessStatusCode);
    }

    [Fact]
    public void EndpointFileIsReadOnlyWhileItsProcessRuns()
    {
        new Core.Hosting.ApiEndpoint("http://127.0.0.1:5123", "K", Environment.ProcessId).Write(_dir);
        Assert.Equal("http://127.0.0.1:5123", Core.Hosting.ApiEndpoint.Read(_dir)?.Url);
        Core.Hosting.ApiEndpoint.Delete(_dir, Environment.ProcessId + 1); // another instance's file is left alone
        Assert.NotNull(Core.Hosting.ApiEndpoint.Read(_dir));
        Core.Hosting.ApiEndpoint.Delete(_dir, Environment.ProcessId);
        Assert.Null(Core.Hosting.ApiEndpoint.Read(_dir));

        new Core.Hosting.ApiEndpoint("http://127.0.0.1:5123", "K", int.MaxValue).Write(_dir); // a crashed instance
        Assert.Null(Core.Hosting.ApiEndpoint.Read(_dir));
    }

    private string Write(string name, string json)
    {
        string path = Path.Combine(_dir, name);
        File.WriteAllText(path, json);
        return path;
    }

    private async Task ImportSettings(HostCommands cli)
    {
        Assert.Equal(0, await cli.PutSetting("connections", Write("app.json", $$"""
            { "name": "app-a", "tenantId": "contoso.onmicrosoft.com", "clientId": "{{Guid.NewGuid()}}",
              "certificateThumbprint": "{{new string('C', 40)}}", "licenses": "upTo5000" }
            """), null));
        Assert.Equal(0, await cli.PutSetting("templates", Write("tpl.json", $$"""
            { "name": "finance", "rules": [ { "old": "{{Host}}/Document", "new": "{{Host}}/sites/finance" } ] }
            """), null));
        Assert.Equal(0, await cli.PutSetting("profiles", Write("finance.profile.json", $$"""
            {
              // as exported from the GUI
              "baseHost": "{{Host}}",
              "templates": [ "finance" ],
              "targets": [ { "siteUrl": "{{Host}}/sites/finance", "library": "Documents" } ],
              "connections": [ "app-a" ],
              "retry": { "filePausesMinutes": [ 0.0005 ], "requestAttempts": 2, "fileAttempts": 2 }
            }
            """), null));
    }

    [Fact]
    public async Task FullCycleFromTheCli()
    {
        var a = _graph.AddFile(_library, "Fin/a.xlsx", TestXlsx.WithHyperlink(OldUrl));
        var b = _graph.AddFile(_library, "Fin/b.xlsx", TestXlsx.WithHyperlink(OldUrl));
        _graph.AddFile(_library, "Fin/ok.xlsx", TestXlsx.WithHyperlink(NewUrl));
        var cli = Cli();
        await ImportSettings(cli);
        Assert.Equal(0, await cli.ValidateProfile("finance"));

        bool broken = true;
        _graph.Faults.Add(r => broken && r.ItemName == "b.xlsx" && r.Path.Contains("/content", StringComparison.Ordinal) && r.Method == "GET"
            ? FakeGraph.Error(HttpStatusCode.ServiceUnavailable, "serverBusy", 0)
            : null);

        Assert.Equal(ExitCodes.Attention, await cli.StartRun("apply", "finance", wait: true, yes: true, CancellationToken.None));
        Assert.Contains("failed:transient", _out.ToString());
        Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(a.Content));

        string reports = Path.Combine(_dir, "reports");
        Assert.Equal(0, await cli.Export(1, reports));
        Assert.Contains("b.xlsx", File.ReadAllText(Path.Combine(reports, "attention.csv")));
        Assert.Contains("retry", File.ReadAllText(Path.Combine(reports, "attention.csv")));

        broken = false;
        Assert.Equal(0, await cli.Retry(1, new RetryArgs([], [], [], null, [], null, null, false), wait: true, yes: true, CancellationToken.None));
        Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(b.Content));

        // verify over the files written by the retry run
        Assert.Equal(0, await cli.VerifyRun(2, wait: true, CancellationToken.None));
        var store = _factory.Services.GetRequiredService<StateStore>();
        Assert.All(store.QueryFiles(3, new FileFilter(), 0, 100), f => Assert.Equal(FileStatuses.Verified, f.StatusCode));

        Assert.Equal(0, await cli.Runs(10));
        Assert.Contains("verify", _out.ToString());
    }

    [Fact]
    public async Task InvalidProfileIsReportedWithIssues()
    {
        var cli = Cli();
        string file = Write("bad.json", """{ "rules": [ { "old": "relative/path", "new": "x" } ], "connections": [ "missing" ] }""");
        Assert.Equal(ExitCodes.InvalidInput, await cli.PutSetting("profiles", file, "bad"));
        Assert.Contains("ПОМИЛКА", _out.ToString());

        var ex = await Assert.ThrowsAsync<HostException>(() => cli.StartRun("scan", "bad", false, true, CancellationToken.None));
        Assert.Equal(HttpStatusCode.BadRequest, ex.Status);
        Assert.Contains(ex.Issues, i => i.Contains("missing"));

        var notFound = await Assert.ThrowsAsync<HostException>(() => cli.Status(999));
        Assert.Equal(HttpStatusCode.NotFound, notFound.Status);
    }

    [Fact]
    public async Task OnlyOneRunAtATimeAndPauseStop()
    {
        for (int i = 0; i < 10; i++)
            _graph.AddFile(_library, $"f{i}.xlsx", TestXlsx.WithHyperlink(OldUrl));
        var cli = Cli();
        await ImportSettings(cli);
        _graph.Faults.Add(r => { Thread.Sleep(100); return null; });

        Assert.Equal(0, await cli.StartRun("scan", "finance", wait: false, yes: false, CancellationToken.None));
        var busy = await Assert.ThrowsAsync<HostException>(() => cli.StartRun("scan", "finance", false, false, CancellationToken.None));
        Assert.Equal(HttpStatusCode.Conflict, busy.Status);

        Assert.Equal(0, await cli.Control(1, "pause", false, CancellationToken.None));
        Assert.Equal(0, await cli.Control(1, "resume", false, CancellationToken.None));
        Assert.Equal(HostCommands.RunNotFinished, await cli.Control(1, "stop", wait: true, CancellationToken.None));

        // continue the stopped run to the end
        _graph.Faults.Clear();
        Assert.Equal(0, await cli.Control(1, "start", wait: true, CancellationToken.None));
    }

    [Fact]
    public async Task RunIsDeletedFromTheCliOnlyWhenNotRunning()
    {
        for (int i = 0; i < 10; i++)
            _graph.AddFile(_library, $"f{i}.xlsx", TestXlsx.WithHyperlink(OldUrl));
        var cli = Cli();
        await ImportSettings(cli);
        _graph.Faults.Add(r => { Thread.Sleep(100); return null; });
        Assert.Equal(0, await cli.StartRun("scan", "finance", wait: false, yes: false, CancellationToken.None));

        var busy = await Assert.ThrowsAsync<HostException>(() => cli.DeleteRun(1, yes: true, CancellationToken.None));
        Assert.Equal(HttpStatusCode.Conflict, busy.Status);

        Assert.Equal(HostCommands.RunNotFinished, await cli.Control(1, "stop", wait: true, CancellationToken.None));
        Assert.Equal(0, await cli.DeleteRun(1, yes: true, CancellationToken.None));
        Assert.Contains("Прогін 1 видалено", _out.ToString());
        var gone = await Assert.ThrowsAsync<HostException>(() => cli.Status(1));
        Assert.Equal(HttpStatusCode.NotFound, gone.Status);
        var again = await Assert.ThrowsAsync<HostException>(() => cli.DeleteRun(1, yes: true, CancellationToken.None));
        Assert.Equal(HttpStatusCode.NotFound, again.Status);
    }

    [Fact]
    public async Task InterruptedRunIsContinuedWhenTheHostStarts()
    {
        for (int i = 0; i < 3; i++)
            _graph.AddFile(_library, $"f{i}.xlsx", TestXlsx.WithHyperlink(OldUrl));
        await ImportSettings(Cli());
        var service = _factory.Services.GetRequiredService<RunService>();
        long id = service.CreateRun(RunModes.Apply, "finance");
        // The host died while the run was going: the state says "running".
        service.Store.SetRunStatus(id, RunStatuses.Running, DateTimeOffset.UtcNow);
        await _factory.DisposeAsync();

        _factory = CreateFactory(continueInterrupted: true);
        var cli = Cli();
        Assert.Equal(0, await cli.Wait(id, CancellationToken.None));
        Assert.All(_graph.Log.Where(r => r.Method == "PUT"), r => Assert.Equal("app-a", r.App));
    }

    [Fact]
    public async Task ExportsForTheGui()
    {
        _graph.AddFile(_library, "Fin/a.xlsx", TestXlsx.WithHyperlink(OldUrl));
        _graph.AddFile(_library, "Fin/b.xlsx", TestXlsx.WithHyperlink(NewUrl));
        var cli = Cli();
        await ImportSettings(cli);
        Assert.Equal(0, await cli.StartRun("scan", "finance", wait: true, yes: true, CancellationToken.None));
        var http = Http();

        var xlsx = await http.GetAsync("api/runs/1/files/export?format=xlsx&status=planned");
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx.Content.Headers.ContentType!.MediaType);
        var rows = OpenXml.SimpleXlsxReader.ReadFirstSheet(await xlsx.Content.ReadAsStreamAsync());
        Assert.Equal(2, rows.Count); // header + the one planned file
        Assert.Equal("Fin/a.xlsx", rows[1][1]);

        var csv = await http.GetStringAsync("api/runs/1/files/export?format=csv");
        Assert.Equal(3, csv.Split("\r\n", StringSplitOptions.RemoveEmptyEntries).Length);

        var report = await http.GetAsync("api/runs/1/reports/links?format=xlsx");
        Assert.True(report.IsSuccessStatusCode);

        var json = await http.GetAsync("api/profiles/finance?download=true");
        Assert.Equal("attachment", json.Content.Headers.ContentDisposition!.DispositionType);
        Assert.Contains("\"targets\"", await json.Content.ReadAsStringAsync());

    }

    [Fact]
    public void ReadsPathListFromCsv()
    {
        string file = Write("list.csv", "id,path,status\r\n1,Fin/a.xlsx,failed:transient\r\n2,\"Fin/b, c.xlsx\",failed:locked\r\n");
        Assert.Equal(["Fin/a.xlsx", "Fin/b, c.xlsx"], HostCommands.ReadPathList(file));
    }
}
