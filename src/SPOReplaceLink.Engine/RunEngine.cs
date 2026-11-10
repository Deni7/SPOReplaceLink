using System.Net;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Rules;
using SPOReplaceLink.ExcelWorker;
using SPOReplaceLink.OpenXml;
using SPOReplaceLink.Spo;
using SPOReplaceLink.Spo.Graph;
using SPOReplaceLink.State;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Engine;

public sealed class EngineOptions
{
    /// <summary>SEC-3: temporary files of runs; each file is deleted right after processing.</summary>
    public required string TempRoot { get; init; }

    public TimeProvider Time { get; init; } = TimeProvider.System;

    /// <summary>How often the dispatcher looks for files whose pause has ended.</summary>
    public TimeSpan PollInterval { get; init; } = TimeSpan.FromSeconds(1);

    /// <summary>How often metrics are written to run_metrics for charts.</summary>
    public TimeSpan MetricsInterval { get; init; } = TimeSpan.FromMinutes(1);

    /// <summary>Pause between attempts to continue a failed enumeration.</summary>
    public TimeSpan EnumerationRetryDelay { get; init; } = TimeSpan.FromSeconds(30);

    /// <summary>
    /// XLS-4: makes the Excel pool of a run from the profile's process count and timeout (see <see cref="ExcelPool.Factory"/>).
    /// Null — Excel is not used, and xls files stay needs_excel.
    /// </summary>
    public Func<int, TimeSpan, IExcelConverter>? Excel { get; init; }
}

/// <summary>
/// Runs one run to a final status (RTY-7) or until it is stopped. Everything it does is recorded in the state
/// store, so a run interrupted at any point continues from the same place (P-6, STA-3).
/// </summary>
public sealed class RunEngine
{
    private readonly StateStore _store;
    private readonly ISpoAppFactory _apps;
    private readonly ILogger _logger;
    private readonly EngineOptions _options;
    private readonly TimeProvider _time;
    private volatile bool _paused;
    private volatile string? _abortReason;
    private readonly object _breakerSync = new();
    private readonly Queue<bool> _recent = new();
    private int _authFailuresInRow;

    public RunEngine(StateStore store, ISpoAppFactory apps, ILogger logger, EngineOptions options)
    {
        _store = store;
        _apps = apps;
        _logger = logger;
        _options = options;
        _time = options.Time;
        Metrics = new RunMetrics(_time);
    }

    public RunMetrics Metrics { get; }

    public bool IsPaused => _paused;

    /// <summary>STA-5: no new files are taken; files in progress finish.</summary>
    public void Pause() => _paused = true;

    public void Resume() => _paused = false;

    private DateTimeOffset Now => _time.GetUtcNow();

    /// <param name="stop">STA-4: graceful stop — take no new files, finish the current ones, save state.</param>
    /// <param name="abort">Hard stop: current files are interrupted and redone on the next start.</param>
    /// <returns>Final run status.</returns>
    public async Task<string> RunAsync(long runId, CancellationToken stop, CancellationToken abort = default)
    {
        var run = _store.GetRun(runId) ?? throw new InvalidOperationException(L.F("Прогін {0} не знайдено.", runId));
        var config = RunConfig.Parse(run.ConfigSnapshot);
        var profile = config.Profile;
        var validation = profile.ValidateForRun(config.Connections);
        if (!validation.IsValid)
        {
            string reason = L.T("Знімок налаштувань невалідний: ") + string.Join("; ", validation.Issues.Where(i => i.Severity == IssueSeverity.Error));
            _store.SetRunStatus(runId, RunStatuses.Aborted, Now, reason);
            return RunStatuses.Aborted;
        }

        ConfigureBreaker(profile.Retry);
        var ctx = new RunContext(run, config, validation.RuleSet!, new WorkbookProcessor(validation.RuleSet!, profile.Places),
            Path.Combine(_options.TempRoot, $"run-{runId}"));
        Directory.CreateDirectory(ctx.TempDir);

        var connections = profile.Connections
            .Select(n => config.Connections.First(c => string.Equals(c.Name, n, StringComparison.OrdinalIgnoreCase)))
            .ToList();
        var settings = new SpoAppSettings(
            profile.Performance.Downloads + profile.Performance.Uploads,
            profile.Retry.RequestAttempts,
            TimeSpan.FromSeconds(profile.Performance.RequestTimeoutSeconds),
            TimeSpan.FromMinutes(profile.Retry.BlockedMinutes),
            profile.Performance.BudgetPercent / 100.0);
        try
        {
            ctx.Pool = AppPool.Create(_apps, connections, settings, profile.Pool.TenantCeilingPercent, _time);
        }
        catch (SpoException ex)
        {
            return Abort(runId, L.F("Підключення: {0}", ex.Message));
        }
        Metrics.AttachPool(ctx.Pool);
        await using var excel = _options.Excel is { } makeExcel && profile.Performance.ExcelProcesses > 0
            ? makeExcel(profile.Performance.ExcelProcesses, TimeSpan.FromMinutes(profile.Performance.ExcelTimeoutMinutes))
            : null;
        if (excel is not null)
        {
            ctx.Excel = new ExcelStage(ctx.Processor, excel);
            Metrics.AttachExcel(excel);
        }

        _store.SetRunStatus(runId, RunStatuses.Running, Now);
        int recovered = _store.RecoverInProgress(runId, Now);
        if (recovered > 0)
            _logger.LogInformation(L.T("Прогін {Run}: {Count} файлів після перерваної обробки повернуто в чергу"), runId, recovered);
        long found = _store.CountFiles(runId, new FileFilter());
        Metrics.SetFound(found);
        Metrics.SetFinishedBefore(found - _store.CountNonFinal(runId));

        try
        {
            await PrepareTargetsAsync(ctx, abort).ConfigureAwait(false);
        }
        catch (SpoException ex)
        {
            return Abort(runId, L.F("Підготовка цілей: {0}", ex.Message));
        }

        using var enumerationStop = CancellationTokenSource.CreateLinkedTokenSource(stop, abort);
        var enumeration = run.Enumerate && _store.GetTargets(runId).Any(t => !t.Done)
            ? Task.Run(() => EnumerateAsync(ctx, enumerationStop.Token), CancellationToken.None)
            : Task.CompletedTask;

        string status;
        try
        {
            status = await DispatchAsync(ctx, enumeration, stop, abort).ConfigureAwait(false);
        }
        finally
        {
            enumerationStop.Cancel();
            try
            {
                await enumeration.ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is OperationCanceledException || enumeration.IsFaulted)
            {
                // Already turned into the stop reason by the dispatcher.
            }
            TryDeleteDirectory(ctx.TempDir);
            _store.AddMetrics(runId, JsonSerializer.Serialize(Metrics.Snapshot(), Profile.JsonOptions), Now);
        }
        return status;
    }

    private string Abort(long runId, string reason)
    {
        _logger.LogError(L.T("Прогін {Run} зупинено: {Reason}"), runId, reason);
        _store.SetRunStatus(runId, RunStatuses.Aborted, Now, reason);
        // RU spent on resolving targets count too (pilot report).
        _store.AddMetrics(runId, JsonSerializer.Serialize(Metrics.Snapshot(), Profile.JsonOptions), Now);
        return RunStatuses.Aborted;
    }

    private sealed class RunContext(RunRow run, RunConfig config, RuleSet rules, WorkbookProcessor processor, string tempDir)
    {
        public RunRow Run { get; } = run;
        public RunConfig Config { get; } = config;
        public Profile Profile => Config.Profile;
        public RuleSet Rules { get; } = rules;
        public WorkbookProcessor Processor { get; } = processor;
        public string TempDir { get; } = tempDir;
        public AppPool Pool { get; set; } = null!;
        public ExcelStage? Excel { get; set; }
        public SemaphoreSlim Downloads { get; } = new(config.Profile.Performance.Downloads);
        public SemaphoreSlim Processors { get; } = new(config.Profile.Performance.Processors);
        public SemaphoreSlim Uploads { get; } = new(config.Profile.Performance.Uploads);
        public volatile bool EnumerationDone;
        public string? EnumerationError;
        public Dictionary<int, TargetRow> Targets { get; } = [];
    }

    // ------------------------------------------------------------------ targets

    /// <summary>Resolves site, library and folder of each target and checks that every app of the pool can reach them (THR-11).</summary>
    private async Task PrepareTargetsAsync(RunContext ctx, CancellationToken ct)
    {
        var rows = _store.GetTargets(ctx.Run.Id);
        for (int i = 0; i < ctx.Profile.Targets.Count && i < rows.Count; i++)
        {
            var target = ctx.Profile.Targets[i];
            var row = rows[i];
            if (row.DriveId is null)
            {
                var (drive, rootItem, rootPath) = await TargetResolver.ResolveAsync(ctx.Pool.Primary.Client, target, ct).ConfigureAwait(false);
                _store.SetTargetDrive(ctx.Run.Id, i, drive.Id, drive.WebUrl, rootItem, rootPath);
            }
        }
        foreach (var row in _store.GetTargets(ctx.Run.Id))
            ctx.Targets[row.Index] = row;

        foreach (var app in ctx.Pool.Apps)
        {
            foreach (var (target, row) in ctx.Profile.Targets.Zip(ctx.Targets.Values.OrderBy(t => t.Index)))
            {
                try
                {
                    var (drive, _, _) = await TargetResolver.ResolveAsync(app.Client, target, ct).ConfigureAwait(false);
                    if (drive.Id != row.DriveId)
                        throw new SpoException(ErrorClass.Internal, L.F("бібліотека визначилась інакше ({0} замість {1}).", drive.Id, row.DriveId));
                }
                catch (SpoException ex)
                {
                    throw new SpoException(ex.ErrorClass, L.F("застосунок «{0}» не має доступу до цілі {1}: {2}", app.Name, target, ex.Message), ex.Status, ex.Code, ex.RequestId, ex);
                }
            }
        }
    }

    // ------------------------------------------------------------------ enumeration (ENU-1, ENU-3)

    private async Task EnumerateAsync(RunContext ctx, CancellationToken ct)
    {
        var folders = new Dictionary<(string Drive, string Item), FolderRow>();
        foreach (var f in _store.GetFolders(ctx.Run.Id))
            folders[(f.DriveId, f.ItemId)] = f;

        foreach (var target in _store.GetTargets(ctx.Run.Id).Where(t => !t.Done))
        {
            int failures = 0;
            string? link = target.NextLink;
            // All pages were saved before an interruption; only the final step is left.
            bool done = link is null && target.DeltaLink is not null;
            while (!done)
            {
                ct.ThrowIfCancellationRequested();
                DeltaPage page;
                try
                {
                    page = await ctx.Pool.Primary.Client.GetDeltaPageAsync(target.DriveId!, link, ct).ConfigureAwait(false);
                    failures = 0;
                }
                catch (SpoException ex) when (ex is not SpoBlockedException && ex.ErrorClass == ErrorClass.Transient && ++failures <= 5)
                {
                    _logger.LogWarning(L.T("Перелік цілі {Target}: {Message}; продовження з тієї ж позиції"), target.Index, ex.Message);
                    await Task.Delay(_options.EnumerationRetryDelay, _time, ct).ConfigureAwait(false);
                    continue;
                }
                catch (SpoException ex)
                {
                    ctx.EnumerationError = L.F("Перелік файлів цілі {0}: {1}", target.Index + 1, ex.Message);
                    _abortReason ??= ctx.EnumerationError;
                    return;
                }

                var newFolders = page.Items
                    .Where(i => (i.IsFolder || i.IsRoot) && !i.Deleted)
                    .Select(i => new FolderRow(target.DriveId!, i.Id, i.ParentId, i.Name, i.IsRoot))
                    .ToList();
                foreach (var f in newFolders)
                    folders[(f.DriveId, f.ItemId)] = f;

                var files = new List<FileRow>();
                foreach (var item in page.Items.Where(i => i.IsFile && !i.Deleted && ctx.Profile.Filters.Accepts(i.Name)))
                {
                    string? dir = FolderPath(folders, target.DriveId!, item.ParentId);
                    var row = new FileRow
                    {
                        RunId = ctx.Run.Id,
                        TargetIndex = target.Index,
                        DriveId = target.DriveId!,
                        ItemId = item.Id,
                        ParentId = item.ParentId,
                        Name = item.Name,
                        Size = item.Size,
                        ETag = item.ETag,
                        Modified = item.Modified,
                        ModifiedBy = item.ModifiedBy,
                        Status = FileStatuses.WaitingPath,
                    };
                    if (dir is not null)
                    {
                        var placed = Place(ctx, target, row, Combine(dir, item.Name));
                        if (placed is null)
                            continue;
                        row = placed;
                    }
                    files.Add(row);
                }

                done = page.NextLink is null;
                _store.SaveDeltaPage(ctx.Run.Id, target.Index, newFolders, files, page.NextLink, page.DeltaLink, done: false, Now);
                Metrics.SetFound(_store.CountFiles(ctx.Run.Id, new FileFilter()));
                link = page.NextLink;
            }

            // Files that came before their folder.
            foreach (var row in _store.GetWaitingPath(ctx.Run.Id).Where(r => r.TargetIndex == target.Index))
            {
                string? dir = FolderPath(folders, row.DriveId, row.ParentId);
                var placed = dir is null ? null : Place(ctx, target, row, Combine(dir, row.Name));
                if (dir is null)
                    _store.ResolvePath(row.Id, null, FileStatuses.Failed("internal"), L.T("Не вдалося визначити папку файлу з delta."), Now);
                else if (placed is null)
                    _store.DeleteFile(row.Id); // outside the target folder, same as files placed during the listing
                else
                    _store.ResolvePath(row.Id, placed.Path, placed.Status, placed.Error, Now);
            }
            _store.SaveDeltaPage(ctx.Run.Id, target.Index, [], [], null, null, done: true, Now);
            _logger.LogInformation(L.T("Перелік цілі {Target} завершено"), target.Index + 1);
        }
        ctx.EnumerationDone = true;
    }

    /// <summary>Sets the path and the filter verdict; null if the file is outside the target folder.</summary>
    private static FileRow? Place(RunContext ctx, TargetRow target, FileRow row, string path)
    {
        string root = target.RootPath ?? string.Empty;
        string relative;
        if (root.Length == 0)
            relative = path;
        else if (path.StartsWith(root + "/", StringComparison.OrdinalIgnoreCase))
            relative = path[(root.Length + 1)..];
        else
            return null;

        var verdict = ctx.Profile.Filters.Check(relative, row.Size, row.Modified);
        return row with
        {
            Path = path,
            Status = verdict switch
            {
                FilterVerdict.Accepted => FileStatuses.Pending,
                FilterVerdict.TooLarge => FileStatuses.SkippedTooLarge,
                _ => FileStatuses.SkippedFiltered,
            },
            Error = verdict is FilterVerdict.Accepted ? null : L.F("Фільтр: {0}", verdict),
        };
    }

    private static string? FolderPath(Dictionary<(string, string), FolderRow> folders, string drive, string? parentId)
    {
        var parts = new List<string>();
        string? id = parentId;
        for (int depth = 0; depth < 400; depth++)
        {
            if (id is null || !folders.TryGetValue((drive, id), out var folder))
                return null;
            if (folder.IsRoot)
            {
                parts.Reverse();
                return string.Join('/', parts);
            }
            parts.Add(folder.Name);
            id = folder.ParentId;
        }
        return null;
    }

    private static string Combine(string dir, string name) => dir.Length == 0 ? name : dir + "/" + name;

    // ------------------------------------------------------------------ dispatcher

    private async Task<string> DispatchAsync(RunContext ctx, Task enumeration, CancellationToken stop, CancellationToken abort)
    {
        var perf = ctx.Profile.Performance;
        int workers = perf.Downloads + perf.Uploads;
        var inFlight = new List<Task>();
        long runId = ctx.Run.Id;
        var nextMetrics = Now + _options.MetricsInterval;
        bool scheduledPause = false;
        bool dailyPause = false;

        try
        {
        while (true)
        {
            inFlight.RemoveAll(t => t.IsCompleted);
            if (abort.IsCancellationRequested)
            {
                await Task.WhenAll(inFlight).ConfigureAwait(false);
                break;
            }
            if (stop.IsCancellationRequested || _abortReason is not null)
            {
                await Task.WhenAll(inFlight).ConfigureAwait(false);
                break;
            }
            if (ctx.Pool.AllBlocked)
                _abortReason ??= new SpoBlockedException(ctx.Pool.Primary.Name, TimeSpan.FromMinutes(ctx.Profile.Retry.BlockedMinutes)).Message;
            if (enumeration.IsFaulted)
                _abortReason ??= L.T("Перелік файлів: ") + (enumeration.Exception?.GetBaseException().Message ?? L.T("невідома помилка"));

            if (Now >= nextMetrics)
            {
                _store.AddMetrics(runId, JsonSerializer.Serialize(Metrics.Snapshot(), Profile.JsonOptions), Now);
                nextMetrics = Now + _options.MetricsInterval;
            }

            // THR-5: budget of the current time window, or wait for the next window.
            int? budget = perf.BudgetAt(Now);
            if (budget is null)
            {
                if (!scheduledPause)
                    _logger.LogInformation(L.T("Прогін {Run}: поза вікнами роботи, очікування"), runId);
                scheduledPause = true;
                await DelayOrAny(inFlight, TimeSpan.FromSeconds(30), stop).ConfigureAwait(false);
                continue;
            }
            scheduledPause = false;
            ctx.Pool.SetBudgetScale(budget.Value / 100.0);

            // THR-6: with the 24-hour budget spent, files would only time out waiting for it; wait here instead.
            if (ctx.Pool.AllDailyExhausted)
            {
                if (!dailyPause)
                    _logger.LogWarning(L.T("Прогін {Run}: добовий бюджет RU вичерпано, пауза до відновлення"), runId);
                dailyPause = true;
                await DelayOrAny(inFlight, TimeSpan.FromMinutes(1), stop).ConfigureAwait(false);
                continue;
            }
            dailyPause = false;

            int free = workers - inFlight.Count;
            IReadOnlyList<FileRow> claimed = [];
            if (!_paused && free > 0)
                claimed = _store.ClaimReady(runId, free, Now);
            foreach (var row in claimed)
                inFlight.Add(Task.Run(() => GuardedAsync(ctx, row, stop, abort), CancellationToken.None));

            if (claimed.Count > 0)
                continue;

            if (inFlight.Count == 0 && !_paused && (enumeration.IsCompleted || ctx.EnumerationDone))
            {
                if (ctx.EnumerationError is not null)
                {
                    _abortReason ??= ctx.EnumerationError;
                    continue;
                }
                if (_store.CountNonFinal(runId) == 0)
                    return Complete(runId);
                // Only files waiting for their pause (level 2) or the next wave (level 3).
                var next = _store.NextNotBefore(runId);
                var wait = next is { } n ? n - Now : _options.PollInterval;
                await DelayOrAny(inFlight, wait < _options.PollInterval ? _options.PollInterval : wait > TimeSpan.FromSeconds(30) ? TimeSpan.FromSeconds(30) : wait, stop)
                    .ConfigureAwait(false);
                continue;
            }
            await DelayOrAny(inFlight, _options.PollInterval, stop).ConfigureAwait(false);
        }
        }
        finally
        {
            // Never leave workers running after the dispatcher is gone (temp folder, run status).
            await Task.WhenAll(inFlight).ConfigureAwait(false);
        }

        if (abort.IsCancellationRequested)
        {
            _store.SetRunStatus(runId, RunStatuses.Stopped, Now, L.T("Перервано"));
            return RunStatuses.Stopped;
        }
        if (_abortReason is not null)
            return Abort(runId, _abortReason);
        _store.SetRunStatus(runId, RunStatuses.Stopped, Now, L.T("Зупинено"));
        return RunStatuses.Stopped;
    }

    private string Complete(long runId)
    {
        var counts = _store.CountByStatus(runId);
        bool errors = counts.Keys.Any(k => k.StartsWith("failed:", StringComparison.Ordinal));
        string status = errors ? RunStatuses.CompletedWithErrors : RunStatuses.Completed;
        _store.SetRunStatus(runId, status, Now);
        _logger.LogInformation(L.T("Прогін {Run} завершено: {Status}"), runId, status);
        return status;
    }

    private async Task DelayOrAny(List<Task> inFlight, TimeSpan delay, CancellationToken stop)
    {
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(stop);
        var tasks = new List<Task>(inFlight) { Task.Delay(delay, _time, cts.Token) };
        await Task.WhenAny(tasks).ConfigureAwait(false);
        cts.Cancel();
    }

    // ------------------------------------------------------------------ one file

    /// <summary>A worker must never die silently: if even writing its outcome fails, the file goes back to the queue.</summary>
    private async Task GuardedAsync(RunContext ctx, FileRow row, CancellationToken stop, CancellationToken abort)
    {
        try
        {
            await ProcessFileAsync(ctx, row, stop, abort).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, L.T("Файл {File}: не вдалося зберегти результат обробки"), row.Path);
            try
            {
                // Without links and residuals: whatever made the full write fail is not repeated.
                _store.Finish(row.Id, new FileOutcome
                {
                    Status = FileStatuses.Failed("internal"),
                    ErrorClass = "internal",
                    Error = L.T("Не вдалося зберегти результат: ") + ex.Message,
                }, null, null, null, Now);
                Metrics.AddFinished();
            }
            catch (Exception inner)
            {
                // The row stays in_progress and is recovered on the next start (STA-3).
                _abortReason ??= L.T("База стану недоступна: ") + inner.Message;
            }
        }
    }

    private async Task ProcessFileAsync(RunContext ctx, FileRow row, CancellationToken stop, CancellationToken abort)
    {
        var app = ctx.Pool.Acquire();
        var perf = ctx.Profile.Performance;
        string mode = ctx.Run.Mode;
        string ext = Path.GetExtension(row.Name).ToLowerInvariant();
        string input = Path.Combine(ctx.TempDir, $"{row.Id}-{row.Attempts}{ext}");
        string output = Path.Combine(ctx.TempDir, $"{row.Id}-{row.Attempts}-out{ext}");
        string stage = "download";
        Metrics.WorkerStarted();
        using var fileTimeout = CancellationTokenSource.CreateLinkedTokenSource(abort);
        fileTimeout.CancelAfter(TimeSpan.FromMinutes(perf.FileTimeoutMinutes));
        var ct = fileTimeout.Token;
        // Until the upload starts nothing is written, so a stop may interrupt the file and put it back in the queue.
        using var beforeWrite = CancellationTokenSource.CreateLinkedTokenSource(ct, stop);
        var early = beforeWrite.Token;
        bool writing = false;
        try
        {
            // RTY-5: a repeated attempt, verify and a retry run work with the current version, not the one from the listing.
            string? etag = row.ETag;
            if (row.Attempts > 1 || mode == RunModes.Verify || etag is null || ctx.Run.ParentId is not null)
            {
                var meta = await app.Client.GetItemAsync(row.DriveId, row.ItemId, early).ConfigureAwait(false);
                etag = meta.ETag;
            }
            _store.SetStage(row.Id, "downloading", app.Name, Now, etag);
            await ctx.Downloads.WaitAsync(early).ConfigureAwait(false);
            try
            {
                await app.Client.DownloadAsync(row.DriveId, row.ItemId, input, early).ConfigureAwait(false);
            }
            finally
            {
                ctx.Downloads.Release();
            }
            Metrics.AddBytes(new FileInfo(input).Length);
            _store.SetStage(row.Id, "downloaded", app.Name, Now);
            // An earlier attempt may have written the file and lost the answer (timeout, crash): then SharePoint already holds our content.
            bool writtenBefore = row.IntentHash is not null && Hash(input) == row.IntentHash;

            stage = "process";
            await ctx.Processors.WaitAsync(early).ConfigureAwait(false);
            FileResult result;
            string? folder = OldFolder(ctx, row);
            try
            {
                result = await Task.Run(() => ctx.Processor.Process(input, output, folder, row.Path ?? row.Name), early).ConfigureAwait(false);
            }
            finally
            {
                ctx.Processors.Release();
            }
            // XLS-3, XLS-5: xls, and xlsb with old prefixes in binary records, are finished by Excel.
            if (ctx.Excel is not null && ExcelStage.IsNeeded(result))
            {
                stage = "excel";
                _store.SetStage(row.Id, "excel", app.Name, Now);
                result = await ctx.Excel.ProcessAsync(input, output, folder, row.Path ?? row.Name,
                    Path.Combine(ctx.TempDir, $"x{row.Id}-{Guid.NewGuid().ToString("N")[..6]}"), early).ConfigureAwait(false);
            }
            _store.SetStage(row.Id, "processed", app.Name, Now);

            string? versionBefore = null;
            string? etagAfter = null;
            string status = BaseStatus(result, mode);
            if (mode == RunModes.Apply && result.OutputWritten)
            {
                stage = "upload";
                // UPL-2, UPL-4, UPL-6: the file must be exactly the version we downloaded and not checked out.
                var meta = await app.Client.GetItemAsync(row.DriveId, row.ItemId, early).ConfigureAwait(false);
                if (meta.CheckedOutBy is not null)
                    throw new SpoException(ErrorClass.Locked, L.F("Файл узято на редагування (check-out): {0}.", meta.CheckedOutBy));
                if (meta.ETag != etag)
                    throw new SpoException(ErrorClass.Conflict, L.T("Файл змінили під час обробки; буде оброблено нову версію."), HttpStatusCode.PreconditionFailed);
                versionBefore = meta.Version;
                await ctx.Uploads.WaitAsync(early).ConfigureAwait(false);
                try
                {
                    string hash = Hash(output);
                    _store.SetWriteIntent(row.Id, versionBefore, meta.ETag!, hash, Now);
                    writing = true;
                    try
                    {
                        var uploaded = await app.Client.UploadAsync(row.DriveId, row.ItemId, output, meta.ETag!, perf.SimpleUploadMaxMb * 1024L * 1024L, ct)
                            .ConfigureAwait(false);
                        etagAfter = uploaded.ETag;
                        long sent = new FileInfo(output).Length;
                        if (uploaded.Size > 0 && uploaded.Size != sent)
                        {
                            // Pilot question: SharePoint may write library columns into the file. Then the content check
                            // after a lost answer (ReconcileAsync) would not recognise our write.
                            Metrics.AddResizedBySpo();
                            _logger.LogInformation(L.T("Файл {File}: SharePoint зберіг {Stored} байт замість {Sent}"), row.Path, uploaded.Size, sent);
                        }
                    }
                    catch (Exception ex) when (ex is SpoException or HttpRequestException or IOException && !abort.IsCancellationRequested)
                    {
                        // The answer may be lost although the write went through: look at what SharePoint holds now.
                        etagAfter = await ReconcileAsync(app, row, meta.ETag!, hash, input, ct).ConfigureAwait(false);
                        if (etagAfter is null)
                            throw;
                        _logger.LogInformation(L.T("Файл {File}: вивантаження пройшло, хоча відповідь не отримано ({Message})"), row.Path, ex.Message);
                    }
                }
                finally
                {
                    ctx.Uploads.Release();
                }
                status = FileStatuses.Done;
            }
            else if (mode == RunModes.Apply && writtenBefore && result.Status == FileStatus.Unchanged)
            {
                status = FileStatuses.Done;
                etagAfter = etag;
            }

            string code = WithSuffix(status, result);
            _store.Finish(row.Id, new FileOutcome
            {
                Status = status,
                StatusCode = code,
                ErrorClass = result.ErrorClass?.ToString().ToLowerInvariant(),
                Error = result.Error,
                VersionBefore = versionBefore,
                ETagAfter = etagAfter,
                ETag = etag,
                Replacements = result.Replacements,
                Residuals = result.Residuals.Count,
                NeedsExcel = result.NeedsExcel || result.Status == FileStatus.NeedsExcel,
            }, result.Links, result.Residuals,
            new AttemptRow(Now, row.Attempts, stage, null, null, null, code, app.Name), Now);
            if (status == FileStatuses.SkippedEncrypted)
                _logger.LogWarning(L.T("Файл {File}: пропущено — {Message} Після зняття пароля: повторний прогін зі статусом skipped:encrypted"), row.Path, result.Error);
            Metrics.AddReplacements(mode == RunModes.Apply || mode == RunModes.Scan ? result.Links : []);
            Metrics.AddFinished();
            if (result.ErrorClass is { } fileError)
                Metrics.AddError(fileError.ToString().ToLowerInvariant());
            RecordOutcome(failure: result.ErrorClass == ErrorClass.Internal, auth: false);
        }
        catch (SpoBlockedException ex)
        {
            // THR-9: one blocked app does not stop the pool; the file goes to another app. All blocked — the dispatcher stops.
            _logger.LogWarning(L.T("Файл {File}: {Message}"), row.Path, ex.Message);
            _store.Release(row.Id, Now);
        }
        catch (SpoException ex) when (ex.IsNotFound && stage != "process" && !writing)
        {
            _store.Finish(row.Id, new FileOutcome { Status = FileStatuses.SkippedDeleted, Error = L.T("Файл видалено або переміщено після переліку.") }, null, null,
                new AttemptRow(Now, row.Attempts, stage, "file", 404, ex.RequestId, ex.Message, app.Name), Now);
            Metrics.AddFinished();
        }
        catch (SpoException ex)
        {
            Fail(ctx, row, stage, ex.ErrorClass, ex.Message, (int?)ex.Status, ex.RequestId, app.Name);
        }
        catch (ExcelException ex) when (ex.Kind == ExcelErrorKind.Unavailable)
        {
            // Without Excel the file waits for it, as before stage 4; a later run with Excel picks it up.
            string message = L.T("Excel недоступний: ") + ex.Message;
            _store.Finish(row.Id, new FileOutcome { Status = FileStatuses.NeedsExcel, StatusCode = FileStatuses.NeedsExcel, Error = message, NeedsExcel = true },
                null, null, new AttemptRow(Now, row.Attempts, stage, null, null, null, message, app.Name), Now);
            Metrics.AddFinished();
        }
        catch (ExcelException ex)
        {
            Fail(ctx, row, stage, ErrorClass.Excel, ex.Message, null, null, app.Name);
        }
        catch (OperationCanceledException) when (abort.IsCancellationRequested)
        {
            // Left in progress on purpose: recovered and redone on the next start (STA-3).
        }
        catch (OperationCanceledException) when (stop.IsCancellationRequested && !writing)
        {
            _store.Release(row.Id, Now);
        }
        catch (OperationCanceledException)
        {
            // In the Excel stage the file timeout means Excel (waiting for it included), not the network.
            Fail(ctx, row, stage, stage == "excel" ? ErrorClass.Excel : ErrorClass.Transient, L.F("Обробка файлу перевищила {0} хв.", perf.FileTimeoutMinutes), null, null, app.Name);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, L.T("Файл {File}: неочікувана помилка на етапі {Stage}"), row.Path, stage);
            Fail(ctx, row, stage, ErrorClass.Internal, ex.GetType().Name + ": " + ex.Message, null, null, app.Name);
        }
        finally
        {
            app.LeaveFile();
            Metrics.WorkerStopped();
            TryDelete(input);
            TryDelete(output);
        }
    }

    /// <summary>Levels 2 and 3 of FEATURES 9.2: back to the queue with a pause, deferred, or a final failure.</summary>
    private void Fail(RunContext ctx, FileRow row, string stage, ErrorClass errorClass, string message, int? http, string? requestId, string app)
    {
        var retry = ctx.Profile.Retry;
        string cls = errorClass.ToString().ToLowerInvariant();
        var attempt = new AttemptRow(Now, row.Attempts, stage, cls, http, requestId, message, app);
        FileOutcome outcome;
        bool final = true;
        switch (errorClass)
        {
            case ErrorClass.Transient or ErrorClass.Conflict when row.Attempts < retry.FileAttempts:
            // FEATURES 9.1: a hung or crashed Excel gets one more attempt, in a fresh Excel process.
            case ErrorClass.Excel when ExcelFailures(row) == 0:
            {
                double pause = retry.FilePausesMinutes[Math.Min(row.Attempts - 1, retry.FilePausesMinutes.Count - 1)];
                outcome = new FileOutcome { Status = FileStatuses.Pending, ErrorClass = cls, Error = message, NotBefore = Now + TimeSpan.FromMinutes(pause) };
                final = false;
                break;
            }
            case ErrorClass.Locked:
            {
                var since = row.LockedSince ?? Now;
                if (Now - since < TimeSpan.FromMinutes(retry.LockedLimitMinutes))
                {
                    outcome = new FileOutcome
                    {
                        Status = FileStatuses.Deferred,
                        ErrorClass = cls,
                        Error = message,
                        LockedSince = since,
                        NotBefore = Now + TimeSpan.FromMinutes(retry.LockedWaveMinutes),
                    };
                    final = false;
                }
                else
                {
                    outcome = new FileOutcome { Status = FileStatuses.Failed(cls), ErrorClass = cls, Error = message, LockedSince = since };
                }
                break;
            }
            default:
                outcome = new FileOutcome { Status = FileStatuses.Failed(cls), ErrorClass = cls, Error = message };
                break;
        }
        _store.Finish(row.Id, outcome, null, null, attempt, Now);
        Metrics.AddError(cls);
        if (final)
            Metrics.AddFinished();
        _logger.LogWarning(L.T("Файл {File}: {Class} на етапі {Stage} (спроба {Attempt}): {Message}"), row.Path, cls, stage, row.Attempts, message);
        RecordOutcome(failure: errorClass is ErrorClass.Transient or ErrorClass.Auth or ErrorClass.Internal, auth: errorClass == ErrorClass.Auth);
    }

    /// <summary>Earlier attempts of this file that Excel failed; other failures do not use up the Excel retry.</summary>
    private int ExcelFailures(FileRow row) => _store.GetAttempts(row.Id).Count(a => a.ErrorClass == "excel");

    /// <summary>
    /// After a failed or unanswered upload: returns the new eTag if SharePoint holds exactly our content, otherwise null.
    /// If someone else saved the file meanwhile, the failure becomes a conflict and the file is redone from their version.
    /// </summary>
    private static async Task<string?> ReconcileAsync(SpoApp app, FileRow row, string etagBefore, string hash, string scratch, CancellationToken ct)
    {
        Spo.Graph.SpoItem current;
        try
        {
            current = await app.Client.GetItemAsync(row.DriveId, row.ItemId, ct).ConfigureAwait(false);
        }
        catch (SpoException)
        {
            return null;
        }
        if (current.ETag == etagBefore)
            return null; // nothing was written
        await app.Client.DownloadAsync(row.DriveId, row.ItemId, scratch, ct).ConfigureAwait(false);
        if (Hash(scratch) == hash)
            return current.ETag;
        throw new SpoException(ErrorClass.Conflict, L.T("Під час вивантаження файл змінив хтось інший; буде оброблено нову версію."), HttpStatusCode.PreconditionFailed);
    }

    private static string Hash(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
    }

    /// <summary>RSL-5: stops the run when too many of the last files failed, or on mass auth failures.</summary>
    private void RecordOutcome(bool failure, bool auth)
    {
        lock (_breakerSync)
        {
            _authFailuresInRow = auth ? _authFailuresInRow + 1 : 0;
            _recent.Enqueue(failure);
            while (_recent.Count > _breakerWindow)
                _recent.Dequeue();
            if (_authFailuresInRow >= _authLimit)
                _abortReason ??= L.F("Запобіжник: {0} помилок авторизації поспіль (401/403). Перевірте підключення і права Sites.Selected.", _authFailuresInRow);
            else if (_recent.Count >= _breakerWindow && _recent.Count(x => x) > _breakerWindow * _breakerRatio)
                _abortReason ??= L.F("Запобіжник: {0} з останніх {1} файлів завершились помилкою.", _recent.Count(x => x), _breakerWindow);
        }
    }

    private int _breakerWindow = 50;
    private double _breakerRatio = 0.5;
    private int _authLimit = 10;

    /// <summary>Applies the breaker settings of the run (called before the first file).</summary>
    private void ConfigureBreaker(RetryOptions retry)
    {
        _breakerWindow = retry.BreakerWindow;
        _breakerRatio = retry.BreakerFailureRatio;
        _authLimit = retry.AuthFailuresToStop;
    }

    private static string BaseStatus(FileResult result, string mode) => result.Status switch
    {
        FileStatus.Unchanged when mode == RunModes.Verify => FileStatuses.Verified,
        FileStatus.Planned when mode == RunModes.Verify => FileStatuses.ChangesPending,
        FileStatus.Unchanged => FileStatuses.Unchanged,
        FileStatus.Planned or FileStatus.Done => FileStatuses.Planned,
        FileStatus.NeedsExcel => FileStatuses.NeedsExcel,
        FileStatus.SkippedEncrypted => FileStatuses.SkippedEncrypted,
        FileStatus.SkippedTooLarge => FileStatuses.SkippedTooLarge,
        _ => FileStatuses.Failed((result.ErrorClass ?? ErrorClass.Internal).ToString().ToLowerInvariant()),
    };

    private static string WithSuffix(string status, FileResult result)
    {
        if (status is not (FileStatuses.Done or FileStatuses.Planned or FileStatuses.Unchanged or FileStatuses.Verified or FileStatuses.ChangesPending))
            return status;
        if (result.NeedsExcel)
            return status + ":needs_excel";
        return result.Residuals.Count > 0 ? status + ":residual" : status;
    }

    /// <summary>
    /// RUL-6: folder the document lived in before migration — the target's oldLocation plus the relative folder,
    /// or the current folder mapped back through the rules.
    /// </summary>
    private static string? OldFolder(RunContext ctx, FileRow row)
    {
        var target = ctx.Profile.Targets[row.TargetIndex];
        var targetRow = ctx.Targets[row.TargetIndex];
        string path = row.Path ?? row.Name;
        int slash = path.LastIndexOf('/');
        string dir = slash < 0 ? string.Empty : path[..slash];
        if (!string.IsNullOrWhiteSpace(target.OldLocation))
        {
            string root = targetRow.RootPath ?? string.Empty;
            string relative = root.Length > 0 && dir.StartsWith(root, StringComparison.OrdinalIgnoreCase) ? dir[root.Length..].TrimStart('/') : dir;
            return relative.Length == 0 ? target.OldLocation.TrimEnd('/') : target.OldLocation.TrimEnd('/') + "/" + relative;
        }
        string current = (targetRow.DriveUrl ?? string.Empty).TrimEnd('/') + (dir.Length == 0 ? string.Empty : "/" + dir);
        return ctx.Rules.Reverse(current);
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
                Directory.Delete(path, true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
