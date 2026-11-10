using System.Collections.Concurrent;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.ExcelWorker;
using SPOReplaceLink.Spo;

namespace SPOReplaceLink.Engine;

/// <summary>Live counters of a run, kept in memory (plan §6): the dashboard does not query the database for them.</summary>
public sealed class RunMetrics
{
    private long _found;
    private long _finished;
    private long _finishedBefore;
    private long _bytes;
    private long _active;
    private long _resized;
    private readonly ConcurrentDictionary<LinkKind, long> _replacements = new();
    private readonly ConcurrentDictionary<string, long> _errors = new();
    private readonly DateTimeOffset _started;
    private readonly TimeProvider _time;
    private AppPool? _pool;
    private IExcelConverter? _excel;

    public RunMetrics(TimeProvider time)
    {
        _time = time;
        _started = time.GetUtcNow();
    }

    public long Found => Interlocked.Read(ref _found);

    public long Finished => Interlocked.Read(ref _finished);

    internal void AttachPool(AppPool pool) => _pool = pool;

    internal void AttachExcel(IExcelConverter excel) => _excel = excel;

    internal void SetFound(long n) => Interlocked.Exchange(ref _found, n);

    /// <summary>Files finished before this start (a continued run): the count goes on from them, the speed counts only new ones.</summary>
    internal void SetFinishedBefore(long n)
    {
        Interlocked.Exchange(ref _finishedBefore, n);
        Interlocked.Exchange(ref _finished, n);
    }

    internal void AddFinished() => Interlocked.Increment(ref _finished);

    internal void AddBytes(long n) => Interlocked.Add(ref _bytes, n);

    internal void WorkerStarted() => Interlocked.Increment(ref _active);

    internal void WorkerStopped() => Interlocked.Decrement(ref _active);

    internal void AddReplacements(IEnumerable<LinkRecord> links)
    {
        foreach (var group in links.Where(l => l.Action == LinkAction.Replaced && l.Kind != LinkKind.WorkbookPath).GroupBy(l => l.Kind))
            _replacements.AddOrUpdate(group.Key, group.Count(), (_, v) => v + group.Count());
    }

    internal void AddError(string errorClass) => _errors.AddOrUpdate(errorClass, 1, (_, v) => v + 1);

    /// <summary>SharePoint reported a different size for a file than was uploaded: it changed the file while saving it (pilot check).</summary>
    internal void AddResizedBySpo() => Interlocked.Increment(ref _resized);

    public MetricsSnapshot Snapshot()
    {
        var elapsed = _time.GetUtcNow() - _started;
        double minutes = Math.Max(elapsed.TotalMinutes, 1.0 / 60);
        return new MetricsSnapshot(
            _time.GetUtcNow(),
            Found,
            Finished,
            (Finished - Interlocked.Read(ref _finishedBefore)) / minutes,
            Interlocked.Read(ref _bytes) / 1024.0 / 1024.0 / Math.Max(elapsed.TotalSeconds, 1),
            Interlocked.Read(ref _active),
            _replacements.ToDictionary(x => x.Key.ToString(), x => x.Value),
            _errors.ToDictionary(x => x.Key, x => x.Value),
            _pool?.Apps.Select(a => new AppSnapshot(
                a.Name,
                a.Budget.SpentTotal,
                a.Budget.SpentLast24h,
                a.Budget.WindowLoad,
                a.Budget.Capacity,
                Interlocked.Read(ref a.Throttle.Throttled429),
                Interlocked.Read(ref a.Throttle.Throttled503),
                Interlocked.Read(ref a.Throttle.Requests),
                a.Throttle.IsPaused ? a.Throttle.PausedUntil : null,
                a.Throttle.ConcurrencyLimit,
                a.ActiveFiles,
                a.Throttle.IsBlocked)).ToList() ?? [],
            _pool?.TenantBudget.SpentTotal ?? 0,
            _pool?.TenantBudget.WindowLoad ?? 0,
            _excel?.Status is { } x ? new ExcelSnapshot(x.Limit, x.Running, x.Busy, x.Operations, x.Timeouts, x.Crashes, x.Unavailable, x.Dialogs) : null,
            Interlocked.Read(ref _resized),
            _started,
            Interlocked.Read(ref _finishedBefore));
    }
}

public sealed record MetricsSnapshot(
    DateTimeOffset At,
    long Found,
    long Finished,
    double FilesPerMinute,
    double MegabytesPerSecond,
    long ActiveWorkers,
    IReadOnlyDictionary<string, long> Replacements,
    IReadOnlyDictionary<string, long> Errors,
    IReadOnlyList<AppSnapshot> Apps,
    long TenantSpent,
    double TenantUtilization,
    ExcelSnapshot? Excel = null,
    long ResizedBySpo = 0,
    DateTimeOffset? Started = null,
    long FinishedBefore = 0);

/// <summary>XLS-4: the Excel pool of the run.</summary>
public sealed record ExcelSnapshot(int Limit, int Running, int Busy, long Operations, long Timeouts, long Crashes, string? Unavailable, long Dialogs = 0);

/// <summary>THR-8: load of one app of the pool.</summary>
public sealed record AppSnapshot(
    string Name,
    long RuSpent,
    long RuLast24h,
    double MinuteUtilization,
    double MinuteCapacity,
    long Throttled429,
    long Throttled503,
    long Requests,
    DateTimeOffset? PausedUntil,
    int ConcurrencyLimit,
    int ActiveFiles,
    bool Blocked);
