using System.Net;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;
using static SPOReplaceLink.Engine.Tests.EngineEnv;

namespace SPOReplaceLink.Engine.Tests;

/// <summary>Scenarios found in code review of stage 2.</summary>
public class ReviewRegressionTests
{
    [Fact]
    public async Task LostUploadAnswerIsReconciledNotRewritten()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("a.xlsx", OldUrl);
        bool lost = false;
        env.Graph.OnUploaded = (_, _) =>
        {
            if (lost)
                return;
            lost = true;
            throw new HttpRequestException("connection reset after the server stored the file");
        };
        long run = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        var file = env.Files(run).Single();
        Assert.Equal(FileStatuses.Done, file.StatusCode);
        Assert.Equal("1.0", file.VersionBefore);
        Assert.Equal(item.ETag, file.ETagAfter);
        Assert.Equal(1, env.Graph.UploadsOf(item));
        Assert.Equal(1, file.Attempts);
    }

    [Fact]
    public async Task WriteBeforeCrashIsRecognizedOnTheNextAttempt()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("a.xlsx", OldUrl);
        long run = env.Create(RunModes.Apply);
        var active = env.Service.Start(run);
        env.Graph.OnUploaded = (_, _) =>
        {
            env.Graph.OnUploaded = null;
            active.Abort.Cancel();
            throw new OperationCanceledException("host killed");
        };
        Assert.Equal(RunStatuses.Stopped, await active.Completion.WaitAsync(TimeSpan.FromSeconds(30)));
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));

        var file = env.Files(run).Single();
        Assert.Equal(FileStatuses.Done, file.StatusCode);
        Assert.Equal("1.0", file.VersionBefore);
        Assert.Equal(1, env.Graph.UploadsOf(item));
    }

    [Fact]
    public async Task OneBlockedAppDoesNotAbortThePool()
    {
        await using var env = new EngineEnv();
        env.AddConnection("app-b");
        for (int i = 0; i < 8; i++)
            env.AddXlsx($"f{i}.xlsx", OldUrl);
        env.Graph.Faults.Add(r => r.App == "app-b" && !r.Path.Contains("/sites/", StringComparison.Ordinal) && !r.Path.Contains("/delta", StringComparison.Ordinal)
            ? FakeGraph.Error(HttpStatusCode.ServiceUnavailable, "serverBusy", 1)
            : null);
        var p = NewProfile("app-a", "app-b");
        p.Pool.Enabled = true;
        p.Retry.BlockedMinutes = 0.02;
        p.Retry.RequestAttempts = 10;
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run, 90));
        Assert.All(env.Files(run), f => Assert.Equal(FileStatuses.Done, f.StatusCode));
        Assert.True(env.Apps.Created["app-b"].Throttle.IsBlocked);
    }

    [Fact]
    public async Task LockedWavesDoNotUseUpAttempts()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("a.xlsx", OldUrl);
        item.LockedUploads = 4;
        var p = NewProfile();
        p.Retry.FileAttempts = 2;
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        var file = env.Files(run).Single();
        Assert.Equal(FileStatuses.Done, file.StatusCode);
        Assert.Equal(1, file.Attempts);
    }

    [Fact]
    public async Task BrokenListingAbortsInsteadOfHangingOrCompleting()
    {
        await using var env = new EngineEnv();
        env.Graph.DeltaPageSize = 2;
        for (int i = 0; i < 5; i++)
            env.AddXlsx($"f{i}.xlsx", OldUrl);
        env.Graph.Faults.Add(r => r.Path.Contains("/delta", StringComparison.Ordinal) && r.Path.Contains("token=2", StringComparison.Ordinal)
            ? new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("{ \"value\": [", System.Text.Encoding.UTF8, "application/json") }
            : null);
        long run = env.Create(RunModes.Scan);
        Assert.Equal(RunStatuses.Aborted, await env.RunAsync(run));
        Assert.Contains("Перелік файлів", env.Store.GetRun(run)!.StopReason);
    }

    [Fact]
    public async Task SpentDailyBudgetPausesInsteadOfFailingFiles()
    {
        await using var env = new EngineEnv();
        var c = new Connection
        {
            Name = "app-a",
            TenantId = "t",
            ClientId = Guid.NewGuid().ToString(),
            CertificateThumbprint = new string('A', 40),
            AppPerDayOverride = 25, // 20 RU at 80%
        };
        env.Store.PutSetting(SettingKinds.Connection, "app-a", c.ToJson(), DateTimeOffset.UtcNow);
        for (int i = 0; i < 10; i++)
            env.AddXlsx($"f{i}.xlsx", OldUrl);
        var p = NewProfile();
        p.Performance.FileTimeoutMinutes = 1;
        long run = env.Create(RunModes.Apply, p);
        var active = env.Service.Start(run);
        await Task.Delay(3000);
        active.Stop.Cancel();
        Assert.Equal(RunStatuses.Stopped, await active.Completion.WaitAsync(TimeSpan.FromSeconds(90)));
        var files = env.Files(run);
        Assert.DoesNotContain(files, f => f.Status.StartsWith("failed", StringComparison.Ordinal));
        Assert.Contains(files, f => f.Status == FileStatuses.Pending);
        Assert.True(env.Apps.Created["app-a"].Budget.SpentLast24h <= 20);
    }

    [Fact]
    public async Task EditDuringChunkedUploadIsNotOverwritten()
    {
        await using var env = new EngineEnv();
        var item = env.Graph.AddFile(env.Library, "big.xlsx", TestXlsx.WithHyperlink(OldUrl, padBytes: 1_200_000));
        var p = NewProfile();
        p.Performance.SimpleUploadMaxMb = 1;
        bool edited = false;
        env.Graph.Faults.Add(r =>
        {
            if (!edited && r.Path.StartsWith("https://upload.test/", StringComparison.Ordinal) && r.Method == "PUT")
            {
                edited = true;
                item.Content = TestXlsx.WithHyperlink(OldUrl, "saved by user during upload", padBytes: 1_200_000);
                item.Version++;
            }
            return null;
        });
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        var file = env.Files(run).Single();
        Assert.Equal(FileStatuses.Done, file.StatusCode);
        Assert.Equal(2, file.Attempts);
        Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(item.Content));
        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(item.Content));
        using var reader = new StreamReader(zip.GetEntry("xl/sharedStrings.xml")!.Open());
        Assert.Contains("saved by user during upload", reader.ReadToEnd());
    }

    [Fact]
    public async Task RetryWithCurrentProfileRequiresSameTargets()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("a.xlsx", OldUrl);
        long run = env.Create(RunModes.Scan);
        await env.RunAsync(run);
        var p = NewProfile();
        p.Targets[0].Folder = "Fin";
        env.SaveProfile(p);
        Assert.Throws<InvalidOperationException>(() => env.Service.CreateChildRun(run, new FileFilter(), currentProfile: true));
    }
}
