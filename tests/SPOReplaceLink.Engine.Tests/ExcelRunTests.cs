using SPOReplaceLink.ExcelWorker;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;
using static SPOReplaceLink.Engine.Tests.EngineEnv;

namespace SPOReplaceLink.Engine.Tests;

/// <summary>Stage 4 in the engine, with the real pool over workers in --fake mode ("fake xls" = OLE signature + xlsx).</summary>
public class ExcelRunTests
{
    private static readonly byte[] Ole = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    /// <param name="marker">hang / crash: makes the fake worker misbehave on this file.</param>
    private static byte[] FakeXls(string url, string? marker = null)
    {
        byte[] xlsx = TestXlsx.WithHyperlink(url);
        if (marker is not null)
        {
            using var ms = new MemoryStream();
            ms.Write(xlsx);
            using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
                zip.CreateEntry("fake-" + marker);
            xlsx = ms.ToArray();
        }
        return [.. Ole, .. xlsx];
    }

    /// <summary>The profile's timeout is at least a minute; tests use seconds.</summary>
    private static Func<string, Func<int, TimeSpan, IExcelConverter>> FakeExcel(List<ExcelPool> pools, double timeoutSeconds = 30, string? worker = null) =>
        dir => (processes, _) =>
        {
            var pool = new ExcelPool(new ExcelPoolOptions(processes, TimeSpan.FromSeconds(timeoutSeconds), dir)
            {
                Fake = true,
                WorkerPath = worker ?? ExcelPoolOptions.DefaultWorkerPath,
            });
            lock (pools)
                pools.Add(pool);
            return pool;
        };

    [Fact]
    public async Task XlsGoesThroughExcelAndIsUploadedInItsFormat()
    {
        var pools = new List<ExcelPool>();
        await using var env = new EngineEnv(FakeExcel(pools));
        var xls = env.Graph.AddFile(env.Library, "Fin/old.xls", FakeXls(OldUrl));
        env.AddXlsx("Fin/a.xlsx", OldUrl);

        long apply = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(apply));
        Assert.Equal(FileStatuses.Done, env.StatusOf(apply, "Fin/old.xls"));
        Assert.True(xls.Content.AsSpan().StartsWith(Ole));
        Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(xls.Content[Ole.Length..]));
        Assert.Contains(env.Store.GetLinks(env.Files(apply).Single(f => f.Name == "old.xls").Id), l => l.NewValue == NewUrl);

        // Excel processes live only while the run does.
        Assert.Equal(0, Assert.Single(pools).Status.Running);

        long verify = env.Create(RunModes.Verify);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(verify));
        Assert.Equal(FileStatuses.Verified, env.StatusOf(verify, "Fin/old.xls"));
    }

    [Fact]
    public async Task SharedWorkbookSavedWithoutSharingIsDoneWithAWarning()
    {
        var pools = new List<ExcelPool>();
        await using var env = new EngineEnv(FakeExcel(pools));
        var xls = env.Graph.AddFile(env.Library, "Fin/shared.xls", FakeXls(OldUrl, "shared"));

        long apply = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(apply));
        var file = env.Files(apply).Single();
        Assert.Equal(FileStatuses.Done, file.StatusCode);
        Assert.Equal(Protocol.UnsharedWarning, file.Error);
        Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(xls.Content[Ole.Length..]));
    }

    [Fact]
    public async Task HungExcelGetsOneMoreAttemptAndDoesNotStopTheRun()
    {
        var pools = new List<ExcelPool>();
        await using var env = new EngineEnv(FakeExcel(pools, timeoutSeconds: 2));
        env.Graph.AddFile(env.Library, "Fin/hang.xls", FakeXls(OldUrl, "hang"));
        env.Graph.AddFile(env.Library, "Fin/b.xls", FakeXls(OldUrl));
        for (int i = 0; i < 5; i++)
            env.AddXlsx($"Fin/{i}.xlsx", OldUrl);

        long id = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.CompletedWithErrors, await env.RunAsync(id, timeoutSeconds: 120));
        var hung = env.Files(id).Single(f => f.Name == "hang.xls");
        Assert.Equal("failed:excel", hung.StatusCode);
        Assert.Equal(2, env.Store.GetAttempts(hung.Id).Count(a => a.ErrorClass == "excel"));
        Assert.All(env.Files(id).Where(f => f.Name != "hang.xls"), f => Assert.Equal(FileStatuses.Done, f.StatusCode));
        var status = Assert.Single(pools).Status;
        Assert.Equal(2, status.Timeouts);
        Assert.Equal(0, status.Running);
    }

    [Fact]
    public async Task WithoutExcelXlsWaitsAsNeedsExcel()
    {
        var pools = new List<ExcelPool>();
        await using var env = new EngineEnv(FakeExcel(pools, worker: Path.Combine(Path.GetTempPath(), "no-such-worker.exe")));
        env.Graph.AddFile(env.Library, "Fin/a.xls", FakeXls(OldUrl));
        env.Graph.AddFile(env.Library, "Fin/b.xls", FakeXls(OldUrl));
        env.AddXlsx("Fin/c.xlsx", OldUrl);

        long id = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(id));
        var xls = env.Files(id).Where(f => f.Name.EndsWith(".xls", StringComparison.Ordinal)).ToList();
        Assert.All(xls, f => Assert.Equal(FileStatuses.NeedsExcel, f.StatusCode));
        Assert.All(xls, f => Assert.Contains("Excel недоступний", f.Error, StringComparison.Ordinal));
        Assert.Equal(FileStatuses.Done, env.StatusOf(id, "Fin/c.xlsx"));
    }

    [Fact]
    public async Task ProfileWithoutExcelProcessesKeepsNeedsExcel()
    {
        var pools = new List<ExcelPool>();
        await using var env = new EngineEnv(FakeExcel(pools));
        env.Graph.AddFile(env.Library, "Fin/a.xls", FakeXls(OldUrl));
        var p = NewProfile();
        p.Performance.ExcelProcesses = 0;

        long id = env.Create(RunModes.Scan, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(id));
        Assert.Equal(FileStatuses.NeedsExcel, env.StatusOf(id, "Fin/a.xls"));
        Assert.Empty(pools);
    }
}
