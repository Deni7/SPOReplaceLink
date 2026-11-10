using System.Diagnostics;

namespace SPOReplaceLink.ExcelWorker.Tests;

/// <summary>XLS-4 with workers in --fake mode: hangs, crashes, restarts and cleanup, without Excel.</summary>
public sealed class ExcelPoolTests : IDisposable
{
    private static readonly byte[] Ole = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "spo-rl-excelpool", Guid.NewGuid().ToString("N"));

    public ExcelPoolTests() => Directory.CreateDirectory(_dir);

    private ExcelPool Pool(int processes = 1, double timeoutSeconds = 30, string? worker = null) =>
        new(new ExcelPoolOptions(processes, TimeSpan.FromSeconds(timeoutSeconds), Path.Combine(_dir, "state"))
        {
            Fake = true,
            WorkerPath = worker ?? ExcelPoolOptions.DefaultWorkerPath,
        });

    private string Book(string name)
    {
        string path = Path.Combine(_dir, name);
        File.WriteAllBytes(path, [.. Ole, 1, 2, 3]);
        return path;
    }

    private string Out() => Path.Combine(_dir, Guid.NewGuid().ToString("N") + ".xlsm");

    internal static bool Exited(int pid)
    {
        try
        {
            using var p = Process.GetProcessById(pid);
            return p.WaitForExit(15_000);
        }
        catch (ArgumentException)
        {
            return true;
        }
    }

    [Fact]
    public async Task ConvertRunsInAWorkerAndReportsTheSourceFormat()
    {
        await using var pool = Pool();
        string output = Out();
        Assert.Equal(XlFormat.Xls, (await pool.ConvertAsync(Book("a.xls"), output, XlFormat.Xlsm, CancellationToken.None)).SourceFormat);
        Assert.Equal([1, 2, 3], await File.ReadAllBytesAsync(output));
        Assert.Single(pool.Processes);
        Assert.Equal(1, pool.Status.Operations);
    }

    [Fact]
    public async Task HungExcelIsKilledAndTheNextFileGetsAFreshProcess()
    {
        await using var pool = Pool(timeoutSeconds: 2);
        await pool.ConvertAsync(Book("a.xls"), Out(), XlFormat.Xlsm, CancellationToken.None);
        int first = pool.Processes.Single().Worker;

        var sw = Stopwatch.StartNew();
        var ex = await Assert.ThrowsAsync<ExcelException>(() => pool.ConvertAsync(Book("hang.xls"), Out(), XlFormat.Xlsm, CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Timeout, ex.Kind);
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(20));
        Assert.True(Exited(first));
        Assert.Empty(pool.Processes);

        await pool.ConvertAsync(Book("b.xls"), Out(), XlFormat.Xlsm, CancellationToken.None);
        Assert.NotEqual(first, pool.Processes.Single().Worker);
        Assert.Equal(1, pool.Status.Timeouts);
    }

    [Fact]
    public async Task CrashedWorkerIsReportedAndReplaced()
    {
        await using var pool = Pool();
        var ex = await Assert.ThrowsAsync<ExcelException>(() => pool.ConvertAsync(Book("crash.xls"), Out(), XlFormat.Xlsm, CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Crashed, ex.Kind);
        Assert.Empty(pool.Processes);
        await pool.ConvertAsync(Book("a.xls"), Out(), XlFormat.Xlsm, CancellationToken.None);
        Assert.Equal(1, pool.Status.Crashes);
    }

    [Fact]
    public async Task PasswordIsAFileErrorAndKeepsTheWorker()
    {
        await using var pool = Pool();
        var ex = await Assert.ThrowsAsync<ExcelException>(() => pool.ConvertAsync(Book("secret.xls"), Out(), XlFormat.Xlsm, CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Password, ex.Kind);
        Assert.Single(pool.Processes);
    }

    [Fact]
    public async Task DialogInExcelRestartsTheWorker()
    {
        await using var pool = Pool();
        await pool.ConvertAsync(Book("a.xls"), Out(), XlFormat.Xlsm, CancellationToken.None);
        int first = pool.Processes.Single().Worker;

        var ex = await Assert.ThrowsAsync<ExcelException>(() => pool.ConvertAsync(Book("prompt.xls"), Out(), XlFormat.Xlsm, CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Password, ex.Kind);
        Assert.Empty(pool.Processes);
        Assert.True(Exited(first));
        Assert.Equal(1, pool.Status.Dialogs);

        await pool.ConvertAsync(Book("b.xls"), Out(), XlFormat.Xlsm, CancellationToken.None);
        Assert.NotEqual(first, pool.Processes.Single().Worker);
    }

    [Fact]
    public async Task MissingWorkerMakesThePoolUnavailableOnce()
    {
        await using var pool = Pool(worker: Path.Combine(_dir, "nope.exe"));
        var first = await Assert.ThrowsAsync<ExcelException>(() => pool.ConvertAsync(Book("a.xls"), Out(), XlFormat.Xlsm, CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Unavailable, first.Kind);
        var second = await Assert.ThrowsAsync<ExcelException>(() => pool.InspectAsync(Book("b.xls"), CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Unavailable, second.Kind);
        Assert.NotNull(pool.Status.Unavailable);
    }

    [Fact]
    public async Task WorkerThatFailsToStartIsACrashNotAMissingExcel()
    {
        // A program that exits at once instead of saying it is ready: one failed start must not switch Excel off for the run.
        await using var pool = Pool(worker: Path.Combine(Environment.SystemDirectory, "where.exe"));
        var ex = await Assert.ThrowsAsync<ExcelException>(() => pool.ConvertAsync(Book("a.xls"), Out(), XlFormat.Xlsm, CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Crashed, ex.Kind);
        Assert.Null(pool.Status.Unavailable);
        Assert.Empty(pool.Processes);
    }

    [Fact]
    public async Task CancelledRequestKillsTheWorker()
    {
        await using var pool = Pool();
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(1));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => pool.ConvertAsync(Book("hang.xls"), Out(), XlFormat.Xlsm, cts.Token));
        Assert.Empty(pool.Processes);
    }

    [Fact]
    public async Task ParallelFilesUseUpToNProcessesAndDisposeLeavesNone()
    {
        var pool = Pool(processes: 2);
        await Task.WhenAll(Enumerable.Range(0, 6).Select(i => pool.ConvertAsync(Book($"f{i}.xls"), Out(), XlFormat.Xlsm, CancellationToken.None)));
        var pids = pool.Processes.Select(p => p.Worker).ToList();
        Assert.InRange(pids.Count, 1, 2);
        await pool.DisposeAsync();
        Assert.All(pids, pid => Assert.True(Exited(pid)));
        Assert.Empty(Directory.GetFiles(Path.Combine(_dir, "state"), "*.pids"));
    }

    [Fact]
    public async Task WorkerIsRecycledAfterTheLimit()
    {
        await using var pool = new ExcelPool(new ExcelPoolOptions(1, TimeSpan.FromSeconds(30), Path.Combine(_dir, "state")) { Fake = true, RecycleAfter = 2 });
        await pool.ConvertAsync(Book("a.xls"), Out(), XlFormat.Xlsm, CancellationToken.None);
        int first = pool.Processes.Single().Worker;
        await pool.ConvertAsync(Book("b.xls"), Out(), XlFormat.Xlsm, CancellationToken.None);
        Assert.Empty(pool.Processes);
        Assert.True(Exited(first));
    }

    [Fact]
    public async Task ProcessesLeftByACrashedHostAreKilledOnTheNextStart()
    {
        string state = Path.Combine(_dir, "state");
        Directory.CreateDirectory(state);
        // A worker of a host that died: its registry file is written but nobody holds it.
        var orphan = Process.Start(new ProcessStartInfo(ExcelPoolOptions.DefaultWorkerPath, "--fake")
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        })!;
        await File.WriteAllTextAsync(Path.Combine(state, "excel-dead.pids"), $"+{orphan.Id} {orphan.StartTime.ToUniversalTime().Ticks}\n");

        // A live pool keeps its own processes.
        await using var live = Pool();
        await live.ConvertAsync(Book("a.xls"), Out(), XlFormat.Xlsm, CancellationToken.None);
        int liveWorker = live.Processes.Single().Worker;
        Assert.True(Exited(orphan.Id));

        await using var next = Pool();
        Assert.False(Process.GetProcessById(liveWorker).HasExited);
        Assert.False(File.Exists(Path.Combine(state, "excel-dead.pids")));
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_dir, true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
