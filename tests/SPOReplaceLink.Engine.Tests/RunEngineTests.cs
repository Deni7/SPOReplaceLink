using System.Net;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;
using static SPOReplaceLink.Engine.Tests.EngineEnv;

namespace SPOReplaceLink.Engine.Tests;

public class RunEngineTests
{
    [Fact]
    public async Task ApplyReplacesUploadsOnlyChangedAndVerifyIsClean()
    {
        await using var env = new EngineEnv();
        var changed = Enumerable.Range(1, 5).Select(i => env.AddXlsx($"Fin/Q{i % 2}/r{i}.xlsx", OldUrl)).ToList();
        var same = env.AddXlsx("Fin/ok.xlsx", NewUrl);
        env.Graph.AddFile(env.Library, "Fin/notes.docx", [1, 2, 3]);
        env.Graph.AddFile(env.Library, "Fin/~$lock.xlsx", [1, 2, 3]);

        long run = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));

        var files = env.Files(run);
        Assert.Equal(6, files.Count); // .docx and ~$ are not taken
        Assert.All(changed, i => Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(i.Content)));
        Assert.All(changed, i => Assert.Equal(1, env.Graph.UploadsOf(i)));
        Assert.Equal(0, env.Graph.UploadsOf(same));
        Assert.Equal(5, files.Count(f => f.StatusCode == FileStatuses.Done));
        Assert.Equal("1.0", files.First(f => f.StatusCode == FileStatuses.Done).VersionBefore);
        Assert.All(files.Where(f => f.Status == FileStatuses.Done), f => Assert.NotNull(f.ETagAfter));

        // verify over the files written by apply
        long verify = env.Service.CreateChildRun(run, new FileFilter { Statuses = [FileStatuses.Done] }, RunModes.Verify);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(verify));
        Assert.All(env.Files(verify), f => Assert.Equal(FileStatuses.Verified, f.StatusCode));
        Assert.Equal(5, env.Files(verify).Count);

        // P-4: a second apply changes and uploads nothing
        long again = env.Service.CreateRun(RunModes.Apply, "test");
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(again));
        Assert.All(env.Files(again), f => Assert.Equal(FileStatuses.Unchanged, f.StatusCode));
        Assert.All(changed, i => Assert.Equal(1, env.Graph.UploadsOf(i)));
    }

    [Fact]
    public async Task ScanDoesNotWrite()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("a.xlsx", OldUrl);
        long run = env.Create(RunModes.Scan);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.Equal(FileStatuses.Planned, env.StatusOf(run, "a.xlsx"));
        Assert.Equal(0, env.Graph.UploadsOf(item));
        Assert.Equal(1, env.Store.GetLinks(env.Files(run)[0].Id).Count(l => l.Action == Core.Processing.LinkAction.Replaced));
    }

    [Fact]
    public async Task FolderTargetFiltersAndLateFolders()
    {
        await using var env = new EngineEnv();
        env.Graph.FilesBeforeFolders = true;
        env.AddXlsx("Fin/2024/a.xlsx", OldUrl);
        env.AddXlsx("Fin/Archive/b.xlsx", OldUrl);
        env.AddXlsx("Other/c.xlsx", OldUrl);
        env.Graph.AddFile(env.Library, "Fin/big.xlsx", TestXlsx.WithHyperlink(OldUrl, padBytes: 2 * 1024 * 1024));
        var p = NewProfile();
        p.Targets[0].Folder = "Fin";
        p.Filters.Exclude = ["Archive/**"];
        p.Filters.MaxFileSizeMb = 1;

        long run = env.Create(RunModes.Scan, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        var files = env.Files(run).ToDictionary(f => f.Path!, f => f.StatusCode);
        Assert.Equal(3, files.Count);
        Assert.Equal(FileStatuses.Planned, files["Fin/2024/a.xlsx"]);
        Assert.Equal(FileStatuses.SkippedFiltered, files["Fin/Archive/b.xlsx"]);
        Assert.Equal(FileStatuses.SkippedTooLarge, files["Fin/big.xlsx"]);
    }

    [Fact]
    public async Task RelativeLinksUseExplicitOldLocation()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("Fin/a.xlsx", "../Analytics/src.xlsx");
        var p = NewProfile();
        p.Targets[0].OldLocation = Host + "/Document";
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.Equal(Host + "/sites/finance/Analytics/src.xlsx", TestXlsx.HyperlinkOf(item.Content));
    }

    [Fact]
    public async Task RelativeLinksUseOldLocationFromReversedRules()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("Fin/a.xlsx", "../Analytics/src.xlsx");
        var p = NewProfile();
        p.Rules = [new RuleEntry { Old = Host + "/Document", New = Host + "/sites/finance/Shared Documents" }];
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.Equal(Host + "/sites/finance/Shared%20Documents/Analytics/src.xlsx", TestXlsx.HyperlinkOf(item.Content));
    }

    [Fact]
    public async Task RelativeLinksWithoutOldLocationAreLeft()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("Fin/a.xlsx", "../Analytics/src.xlsx");
        var p = NewProfile();
        p.Rules = [new RuleEntry { Old = Host + "/Old", New = Host + "/New" }];
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.Equal(FileStatuses.Unchanged, env.StatusOf(run, "Fin/a.xlsx"));
        Assert.Contains(env.Store.GetLinks(env.Files(run)[0].Id), l => l.Action == Core.Processing.LinkAction.Unresolved);
        Assert.Equal(0, env.Graph.UploadsOf(item));
    }
    [Fact]
    public async Task ConflictIsRetriedWithTheNewVersion()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("a.xlsx", OldUrl);
        bool edited = false;
        env.Graph.OnDownloaded = (_, i) =>
        {
            if (edited)
                return;
            edited = true;
            // Someone saves the file between our download and upload.
            i.Content = TestXlsx.WithHyperlink(OldUrl, "edited by user");
            i.Version++;
        };
        long run = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));

        var file = env.Files(run).Single();
        Assert.Equal(FileStatuses.Done, file.StatusCode);
        Assert.Equal(2, file.Attempts);
        Assert.Contains(env.Store.GetAttempts(file.Id), a => a.ErrorClass == "conflict");
        Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(item.Content));
        Assert.Contains("edited by user", System.Text.Encoding.UTF8.GetString(Unzip(item.Content, "xl/sharedStrings.xml")));
    }

    [Fact]
    public async Task LockedFileIsDeferredAndFinishedLater()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("a.xlsx", OldUrl);
        item.LockedUploads = 2;
        long run = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        var file = env.Files(run).Single();
        Assert.Equal(FileStatuses.Done, file.StatusCode);
        Assert.Equal(2, env.Store.GetAttempts(file.Id).Count(a => a.ErrorClass == "locked"));
    }

    [Fact]
    public async Task CheckedOutFileFailsAfterLimit()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("a.xlsx", OldUrl);
        item.CheckedOutBy = "ivan@contoso.com";
        var p = NewProfile();
        p.Retry.LockedLimitMinutes = 0.005;
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.CompletedWithErrors, await env.RunAsync(run));
        Assert.Equal("failed:locked", env.StatusOf(run, "a.xlsx"));
        Assert.Equal(0, env.Graph.UploadsOf(item));
    }

    [Fact]
    public async Task ThrottlingPausesTheWholeAppAndFinishes()
    {
        await using var env = new EngineEnv();
        for (int i = 0; i < 4; i++)
            env.AddXlsx($"f{i}.xlsx", OldUrl);
        int throttled = 0;
        env.Graph.Faults.Add(r => r.Path.Contains("/content", StringComparison.Ordinal) && Interlocked.Increment(ref throttled) <= 2
            ? FakeGraph.Error(HttpStatusCode.TooManyRequests, "activityLimitReached", retryAfterSeconds: 1)
            : null);
        var started = DateTimeOffset.UtcNow;
        long run = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.True(DateTimeOffset.UtcNow - started >= TimeSpan.FromSeconds(1));
        Assert.All(env.Files(run), f => Assert.Equal(FileStatuses.Done, f.StatusCode));
        Assert.Equal(2, env.Apps.Created["app-a"].Throttle.Throttled429);
    }

    [Fact]
    public async Task ExpiredTokenIsRefreshed()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("a.xlsx", OldUrl);
        bool expired = false;
        env.Graph.OnDownloaded = (_, _) =>
        {
            if (!expired)
            {
                expired = true;
                env.Graph.ExpireTokens("app-a");
            }
        };
        long run = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.Equal(FileStatuses.Done, env.StatusOf(run, "a.xlsx"));
        Assert.Equal(1, env.Apps.Tokens["app-a"].Refreshes);
    }

    [Fact]
    public async Task MassAuthFailuresStopTheRun()
    {
        await using var env = new EngineEnv();
        for (int i = 0; i < 6; i++)
            env.AddXlsx($"f{i}.xlsx", OldUrl);
        env.Graph.Faults.Add(r => r.Path.Contains("/content", StringComparison.Ordinal) ? FakeGraph.Error(HttpStatusCode.Forbidden, "accessDenied") : null);
        var p = NewProfile();
        p.Retry.AuthFailuresToStop = 3;
        p.Performance.Downloads = 1;
        p.Performance.Uploads = 1;
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Aborted, await env.RunAsync(run));
        Assert.Contains("авторизації", env.Store.GetRun(run)!.StopReason);
        // Not every file was attempted; the rest waits for the run to be continued.
        Assert.Contains(env.Files(run), f => f.Status == FileStatuses.Pending);
    }

    [Fact]
    public async Task ContinuedRunCountsOnFromTheFilesFinishedBefore()
    {
        await using var env = new EngineEnv();
        for (int i = 0; i < 6; i++)
            env.AddXlsx($"f{i}.xlsx", OldUrl);
        bool denied = true;
        env.Graph.Faults.Add(r => denied && r.ItemName != "f0.xlsx" && r.Path.Contains("/content", StringComparison.Ordinal) ? FakeGraph.Error(HttpStatusCode.Forbidden, "accessDenied") : null);
        var p = NewProfile();
        p.Retry.AuthFailuresToStop = 3;
        p.Performance.Downloads = 1;
        p.Performance.Uploads = 1;
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Aborted, await env.RunAsync(run));
        long before = env.Files(run).Count(f => !FileStatuses.NonFinal.Contains(f.Status));
        Assert.InRange(before, 1, 5);

        denied = false;
        var active = env.Service.Start(run);
        await active.Completion.WaitAsync(TimeSpan.FromSeconds(60));
        // The processed count goes on from the earlier start (the chart no longer drops to zero); the speed counts only new files.
        Assert.DoesNotContain(env.Files(run), f => FileStatuses.NonFinal.Contains(f.Status));
        Assert.Equal(6, active.Engine.Metrics.Finished);
    }

    [Fact]
    public async Task Only503MeansBlockedAndStops()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("a.xlsx", OldUrl);
        var p = NewProfile();
        p.Retry.BlockedMinutes = 0.02;
        p.Retry.RequestAttempts = 20;
        long run = env.Create(RunModes.Apply, p);
        env.Graph.Faults.Add(r => r.Path.Contains("/content", StringComparison.Ordinal) ? FakeGraph.Error(HttpStatusCode.ServiceUnavailable, "serverBusy", 1) : null);
        Assert.Equal(RunStatuses.Aborted, await env.RunAsync(run));
        Assert.Contains("блокування", env.Store.GetRun(run)!.StopReason);
    }

    [Fact]
    public async Task TransientFailureIsRetriedThenRetryRunFixesIt()
    {
        await using var env = new EngineEnv();
        var bad = env.AddXlsx("bad.xlsx", OldUrl);
        env.AddXlsx("good.xlsx", OldUrl);
        var p = NewProfile();
        p.Retry.RequestAttempts = 2;
        p.Retry.FileAttempts = 2;
        bool broken = true;
        env.Graph.Faults.Add(r => broken && r.ItemName == "bad.xlsx" && r.Path.Contains("/content", StringComparison.Ordinal)
            ? FakeGraph.Error(HttpStatusCode.BadGateway, "badGateway")
            : null);
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.CompletedWithErrors, await env.RunAsync(run));
        var failed = env.Files(run).Single(f => f.Path == "bad.xlsx");
        Assert.Equal("failed:transient", failed.StatusCode);
        Assert.Equal(2, failed.Attempts);
        Assert.Equal(2, env.Store.GetAttempts(failed.Id).Count(a => a.ErrorClass == "transient"));
        // Charts: one file failed, and one failed try was followed by another (the last one is the failure itself).
        var (failedAt, retriedAt) = env.Store.ErrorTimeline(run);
        Assert.Single(failedAt);
        Assert.Single(retriedAt);
        Assert.True(retriedAt[0] <= failedAt[0]);

        broken = false;
        long retry = env.Service.CreateChildRun(run, new FileFilter { Statuses = ["failed"] });
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(retry));
        Assert.Equal(FileStatuses.Done, env.Files(retry).Single().StatusCode);
        Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(bad.Content));
    }

    [Fact]
    public async Task IgnoredFilesAreNotRetried()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("bad.xlsx", OldUrl);
        env.AddXlsx("bad2.xlsx", OldUrl);
        var p = NewProfile();
        p.Retry.RequestAttempts = 1;
        p.Retry.FileAttempts = 1;
        env.Graph.Faults.Add(r => r.Path.Contains("/content", StringComparison.Ordinal) && r.Method == "GET" ? FakeGraph.Error(HttpStatusCode.InternalServerError, "x") : null);
        long run = env.Create(RunModes.Apply, p);
        await env.RunAsync(run);
        var first = env.Files(run).First();
        Assert.True(env.Store.Ignore(first.Id, "файл більше не потрібен", DateTimeOffset.UtcNow));
        long retry = env.Service.CreateChildRun(run, new FileFilter { Statuses = ["failed"] });
        Assert.Single(env.Files(retry));
    }

    [Fact]
    public async Task HardStopDuringUploadIsFinishedWithoutSecondUpload()
    {
        await using var env = new EngineEnv();
        var items = Enumerable.Range(0, 4).Select(i => env.AddXlsx($"f{i}.xlsx", OldUrl)).ToList();
        long run = env.Create(RunModes.Apply);

        var active = env.Service.Start(run);
        // The host dies right after SharePoint stored the first upload, before the state was written.
        env.Graph.OnUploaded = (_, _) =>
        {
            env.Graph.OnUploaded = null;
            active.Abort.Cancel();
            throw new OperationCanceledException("host killed");
        };
        Assert.Equal(RunStatuses.Stopped, await active.Completion.WaitAsync(TimeSpan.FromSeconds(30)));
        Assert.Contains(env.Files(run), f => f.Status == FileStatuses.InProgress || f.Status == FileStatuses.Pending);

        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.All(items, i => Assert.Equal(1, env.Graph.UploadsOf(i)));
        Assert.All(items, i => Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(i.Content)));
        // The file written just before the crash is reported as written, not as "unchanged".
        Assert.All(env.Files(run), f => Assert.Equal(FileStatuses.Done, f.StatusCode));
    }

    [Fact]
    public async Task EnumerationContinuesFromSavedPosition()
    {
        await using var env = new EngineEnv();
        env.Graph.DeltaPageSize = 2;
        for (int i = 0; i < 7; i++)
            env.AddXlsx($"f{i}.xlsx", OldUrl);
        int deltaPages = 0;
        int fromStart = 0;
        long run = env.Create(RunModes.Scan);
        RunService.ActiveRun? active = null;
        env.Graph.Faults.Add(r =>
        {
            if (!r.Path.Contains("/delta", StringComparison.Ordinal))
                return null;
            if (!r.Path.Contains("token=", StringComparison.Ordinal))
                Interlocked.Increment(ref fromStart);
            if (Interlocked.Increment(ref deltaPages) == 3)
                active!.Stop.Cancel();
            return null;
        });
        active = env.Service.Start(run);
        Assert.Equal(RunStatuses.Stopped, await active.Completion.WaitAsync(TimeSpan.FromSeconds(30)));
        Assert.NotNull(env.Store.GetTargets(run)[0].NextLink);

        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.Equal(7, env.Files(run).Count);
        // Continued from the saved nextLink: the listing started from the beginning only once.
        Assert.Equal(1, fromStart);
        Assert.InRange(deltaPages, 4, 5);
    }

    [Fact]
    public async Task DeletedAfterListingIsSkipped()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("a.xlsx", OldUrl);
        // The file disappears between listing and download.
        env.Graph.Faults.Add(r => r.ItemName == "a.xlsx" && r.Path.Contains("/content", StringComparison.Ordinal)
            ? FakeGraph.Error(HttpStatusCode.NotFound, "itemNotFound")
            : null);
        long run = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.Equal(FileStatuses.SkippedDeleted, env.StatusOf(run, "a.xlsx"));
    }

    [Fact]
    public async Task LargeFileUsesUploadSessionAndResumes()
    {
        await using var env = new EngineEnv();
        var item = env.Graph.AddFile(env.Library, "big.xlsx", TestXlsx.WithHyperlink(OldUrl, padBytes: 1_500_000));
        var p = NewProfile();
        p.Performance.SimpleUploadMaxMb = 1;
        int chunks = 0;
        env.Graph.Faults.Add(r => r.Path.StartsWith("https://upload.test/", StringComparison.Ordinal) && r.Method == "PUT" && Interlocked.Increment(ref chunks) == 2
            ? throw new HttpRequestException("connection reset")
            : null);
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));
        Assert.Equal(FileStatuses.Done, env.StatusOf(run, "big.xlsx"));
        Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(item.Content));
        Assert.True(chunks >= 5);
    }

    [Fact]
    public async Task OutsideWindowsTheRunWaits()
    {
        await using var env = new EngineEnv();
        var item = env.AddXlsx("a.xlsx", OldUrl);
        var p = NewProfile();
        var now = DateTime.Now;
        // A one-minute window two hours from now.
        p.Performance.Windows = [new ScheduleWindow { From = now.AddHours(2).ToString("HH:mm"), To = now.AddHours(2).AddMinutes(1).ToString("HH:mm") }];
        p.Performance.OnlyInWindows = true;
        long run = env.Create(RunModes.Apply, p);
        var active = env.Service.Start(run);
        await Task.Delay(1500);
        Assert.Equal(0, env.Graph.UploadsOf(item));
        active.Stop.Cancel();
        Assert.Equal(RunStatuses.Stopped, await active.Completion.WaitAsync(TimeSpan.FromSeconds(30)));
    }

    [Fact]
    public async Task InvalidProfileIsRejectedBeforeTheRun()
    {
        await using var env = new EngineEnv();
        var p = NewProfile("app-a", "app-b");
        env.AddConnection("app-b");
        var ex = Assert.Throws<RunValidationException>(() => env.Create(RunModes.Apply, p));
        Assert.Contains(ex.Issues, i => i.Message.Contains("pool.enabled"));
    }

    [Fact]
    public async Task AppWithoutAccessToTargetStopsBeforeAnyWork()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("a.xlsx", OldUrl);
        env.AddConnection("app-b");
        var p = NewProfile("app-a", "app-b");
        p.Pool.Enabled = true;
        env.Graph.Faults.Add(r => r.App == "app-b" && r.Path.Contains("/sites/", StringComparison.Ordinal) ? FakeGraph.Error(HttpStatusCode.Forbidden, "accessDenied") : null);
        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Aborted, await env.RunAsync(run));
        Assert.Contains("app-b", env.Store.GetRun(run)!.StopReason);
        Assert.Empty(env.Files(run));
    }

    [Fact]
    public async Task ReportsExport()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("a.xlsx", OldUrl);
        env.AddXlsx("b.xlsx", @"\\fs01\share\x.xlsx");
        long run = env.Create(RunModes.Scan);
        await env.RunAsync(run);
        string dir = Path.Combine(env.Dir, "reports");
        ReportExporter.ExportAll(env.Store, run, dir);
        string prefixes = File.ReadAllText(Path.Combine(dir, "prefixes.csv"));
        Assert.Contains(Host + "/Document/Fin", prefixes);
        Assert.Contains(@"\\fs01\share\x.xlsx", prefixes);
        Assert.Contains("planned", File.ReadAllText(Path.Combine(dir, "files.csv")));
        Assert.Contains("За статусами", File.ReadAllText(Path.Combine(dir, "summary.txt")));
    }

    private static byte[] Unzip(byte[] zip, string name)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(zip));
        using var s = archive.GetEntry(name)!.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }
}
