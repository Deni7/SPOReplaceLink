using SPOReplaceLink.Core.Config;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;
using static SPOReplaceLink.Engine.Tests.EngineEnv;

namespace SPOReplaceLink.Engine.Tests;

public class RunDeleteTests
{
    [Fact]
    public async Task DeletedRunLeavesNothingAndItsRetryKeepsItsFiles()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("a.xlsx", OldUrl);
        env.AddXlsx("b.xlsx", OldUrl);
        long run = env.Create(RunModes.Scan);
        await env.RunAsync(run);
        var fileIds = env.Files(run).Select(f => f.Id).ToList();
        Assert.Contains(fileIds, id => env.Store.GetLinks(id).Count > 0);
        long retry = env.Service.CreateChildRun(run, new FileFilter { Statuses = [FileStatuses.Planned] });
        await env.RunAsync(retry);
        Assert.NotNull(ProfileInsights.For(env.Store, env.Service.GetProfile(NewProfile().Name!)));

        env.Service.DeleteRun(run);

        Assert.Null(env.Store.GetRun(run));
        Assert.Empty(env.Files(run));
        Assert.All(fileIds, id => Assert.Empty(env.Store.GetLinks(id)));
        Assert.All(fileIds, id => Assert.Empty(env.Store.GetAttempts(id)));
        Assert.Empty(env.Store.GetTargets(run));
        Assert.Empty(env.Store.GetMetrics(run));
        // The estimate of a new run no longer comes from the deleted listing.
        Assert.Null(ProfileInsights.For(env.Store, env.Service.GetProfile(NewProfile().Name!)));

        var child = env.Store.GetRun(retry)!;
        Assert.Null(child.ParentId);
        Assert.Equal(2, env.Files(retry).Count);
        Assert.Throws<KeyNotFoundException>(() => env.Service.DeleteRun(run));
    }

    [Fact]
    public async Task RunningRunCannotBeDeleted()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("a.xlsx", OldUrl);
        using var gate = new ManualResetEventSlim();
        env.Graph.Faults.Add(r =>
        {
            if (r.Path.Contains("/content", StringComparison.Ordinal))
                gate.Wait(TimeSpan.FromSeconds(30));
            return null;
        });
        long run = env.Create(RunModes.Scan);
        var active = env.Service.Start(run);
        for (int i = 0; i < 100 && env.Store.GetRun(run)!.Status != RunStatuses.Running; i++)
            await Task.Delay(50);

        var ex = Assert.Throws<InvalidOperationException>(() => env.Service.DeleteRun(run));
        Assert.Contains("виконується", ex.Message);
        gate.Set();
        await active.Completion.WaitAsync(TimeSpan.FromSeconds(60));
        Assert.NotNull(env.Store.GetRun(run));
    }

    [Fact]
    public async Task RunLeftRunningByAKilledHostCannotBeDeleted()
    {
        await using var env = new EngineEnv();
        long run = env.Create(RunModes.Scan);
        env.Store.SetRunStatus(run, RunStatuses.Running, DateTimeOffset.UtcNow);
        Assert.Throws<InvalidOperationException>(() => env.Service.DeleteRun(run));

        env.Store.SetRunStatus(run, RunStatuses.Stopped, DateTimeOffset.UtcNow, "Зупинено");
        env.Service.DeleteRun(run);
        Assert.Null(env.Store.GetRun(run));
    }
}
