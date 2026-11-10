using System.Net;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;
using static SPOReplaceLink.Engine.Tests.EngineEnv;

namespace SPOReplaceLink.Engine.Tests;

public class PoolTests
{
    private static bool IsWork(FakeGraph.FakeRequest r) =>
        !r.Path.Contains("/sites/", StringComparison.Ordinal) && !r.Path.Contains("/delta", StringComparison.Ordinal);

    [Fact]
    public async Task ThrottlingOfOneAppDoesNotStopTheOthers()
    {
        await using var env = new EngineEnv();
        env.AddConnection("app-b");
        var items = Enumerable.Range(0, 12).Select(i => env.AddXlsx($"f{i}.xlsx", OldUrl)).ToList();
        env.Graph.Faults.Add(r => r.App == "app-b" && IsWork(r) ? FakeGraph.Error(HttpStatusCode.TooManyRequests, "activityLimitReached", 2) : null);
        var p = NewProfile("app-a", "app-b");
        p.Pool.Enabled = true;
        p.Retry.RequestAttempts = 2;

        long run = env.Create(RunModes.Apply, p);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run, 60));

        var files = env.Files(run);
        Assert.All(files, f => Assert.Equal(FileStatuses.Done, f.StatusCode));
        Assert.All(items, i => Assert.Equal(NewUrl, TestXlsx.HyperlinkOf(i.Content)));
        Assert.True(env.Apps.Created["app-b"].Throttle.Throttled429 > 0);
        // Every upload was done by the app that was not throttled.
        Assert.DoesNotContain(env.Graph.Log, r => r.App == "app-b" && r.Method == "PUT");
        Assert.All(files, f => Assert.Equal("app-a", f.App));
    }

    [Fact]
    public async Task PoolStaysUnderTheTenantCeiling()
    {
        await using var env = new EngineEnv();
        // Tenant: 300 RU per 5 minutes; the pool may take 20% = 60 RU, refilled at 0.2 RU/s.
        env.Store.DeleteSetting(SettingKinds.Connection, "app-a", DateTimeOffset.UtcNow);
        env.AddConnection("app-a", tenantOverride: 300);
        env.AddConnection("app-b", tenantOverride: 300);
        for (int i = 0; i < 40; i++)
            env.AddXlsx($"f{i}.xlsx", OldUrl);
        var p = NewProfile("app-a", "app-b");
        p.Pool.Enabled = true;
        p.Pool.TenantCeilingPercent = 20;

        long run = env.Create(RunModes.Apply, p);
        var started = DateTimeOffset.UtcNow;
        var active = env.Service.Start(run);
        await Task.Delay(3000);
        active.Stop.Cancel();
        Assert.Equal(RunStatuses.Stopped, await active.Completion.WaitAsync(TimeSpan.FromSeconds(60)));
        double seconds = (DateTimeOffset.UtcNow - started).TotalSeconds;

        int spent = env.Graph.Log.Sum(env.Graph.CostOf);
        Assert.True(spent <= 60 + (0.2 * seconds) + 2, $"spent {spent} RU in {seconds:0.0} s");
        Assert.Contains(env.Graph.Log, r => r.App == "app-a" && IsWork(r));
        Assert.Contains(env.Graph.Log, r => r.App == "app-b" && IsWork(r));
        // Stopped gracefully: nothing is left half-done, the rest waits for continuation.
        Assert.DoesNotContain(env.Files(run), f => f.Status == FileStatuses.InProgress);
        Assert.Contains(env.Files(run), f => f.Status == FileStatuses.Pending);
    }
}
