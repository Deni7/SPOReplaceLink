using System.Diagnostics;
using System.Text;
using System.Text.Json;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.ExcelWorker;

/// <summary>What the engine needs from Excel; the pool is the real one, tests may put a fake in its place.</summary>
public interface IExcelConverter : IAsyncDisposable
{
    /// <summary>Opens <paramref name="input"/> in Excel and saves it as <paramref name="output"/>; reports the input's format and any warning.</summary>
    Task<WorkerResponse> ConvertAsync(string input, string output, int format, CancellationToken ct);

    /// <summary>The same as <see cref="ConvertAsync"/>, also reporting the links Excel holds when it saves.</summary>
    Task<WorkerResponse> ConvertInspectingAsync(string input, string output, int format, CancellationToken ct);

    /// <summary>The links Excel sees in a workbook (Edit Links, hyperlinks, formulas with text).</summary>
    Task<WorkerResponse> InspectAsync(string input, CancellationToken ct);

    ExcelPoolStatus Status { get; }
}

/// <summary>State of the Excel pool for the dashboard.</summary>
/// <param name="Dialogs">Requests during which Excel showed a dialog (closed by the worker, Excel restarted).</param>
public sealed record ExcelPoolStatus(int Limit, int Running, int Busy, long Operations, long Timeouts, long Crashes, string? Unavailable, long Dialogs = 0);

public sealed record ExcelPoolOptions(int Processes, TimeSpan Timeout, string StateDir)
{
    /// <summary>The worker exe (or dll, then it runs through dotnet).</summary>
    public string WorkerPath { get; init; } = DefaultWorkerPath;

    /// <summary>Workers without Excel (tests of the pool and the engine).</summary>
    public bool Fake { get; init; }

    /// <summary>A worker and its Excel are restarted after this many operations, so Excel never runs for long.</summary>
    public int RecycleAfter { get; init; } = 200;

    public TimeSpan StartTimeout { get; init; } = TimeSpan.FromMinutes(1);

    public static string DefaultWorkerPath
    {
        get
        {
            string exe = Path.Combine(AppContext.BaseDirectory, "SPOReplaceLink.ExcelWorker.exe");
            return File.Exists(exe) ? exe : Path.Combine(AppContext.BaseDirectory, "SPOReplaceLink.ExcelWorker.dll");
        }
    }
}

/// <summary>
/// XLS-4: up to N worker processes, each with its own Excel. A request that gets no answer within the timeout kills
/// the worker and its Excel; the next request starts a fresh one. Every process the pool starts is written to a file
/// in <see cref="ExcelPoolOptions.StateDir"/>, so processes left after a crash of the host are killed on the next start.
/// </summary>
public sealed class ExcelPool : IExcelConverter
{
    private readonly ExcelPoolOptions _options;
    private readonly SemaphoreSlim _slots;
    private readonly object _sync = new();
    private readonly Stack<WorkerProcess> _idle = new();
    private readonly HashSet<WorkerProcess> _all = [];
    private readonly ProcessRegistry _registry;
    private readonly RecentItems? _recent;
    private volatile string? _unavailable;
    private volatile bool _disposed;
    private int _busy;
    private long _operations;
    private long _timeouts;
    private long _crashes;
    private long _dialogs;

    /// <summary>What the engine calls to make the pool of a run: processes and timeout come from the profile.</summary>
    public static Func<int, TimeSpan, IExcelConverter> Factory(string stateDir) =>
        (processes, timeout) => new ExcelPool(new ExcelPoolOptions(processes, timeout, stateDir));

    public ExcelPool(ExcelPoolOptions options)
    {
        if (options.Processes < 1)
            throw new ArgumentOutOfRangeException(nameof(options), L.T("Потрібен хоча б один процес Excel."));
        _options = options;
        _slots = new SemaphoreSlim(options.Processes);
        _registry = new ProcessRegistry(options.StateDir);
        // Excel's temporary copies must not crowd the user's own files out of Recent Items (Quick Access).
        _recent = options.Fake ? null : new RecentItems(options.StateDir);
        _recent?.Sweep();
    }

    public ExcelPoolStatus Status
    {
        get
        {
            lock (_sync)
                return new(_options.Processes, _all.Count, Volatile.Read(ref _busy), Interlocked.Read(ref _operations),
                    Interlocked.Read(ref _timeouts), Interlocked.Read(ref _crashes), _unavailable, Interlocked.Read(ref _dialogs));
        }
    }

    /// <summary>Process ids of the running workers and their Excel instances (tests).</summary>
    internal IReadOnlyList<(int Worker, int? Excel)> Processes
    {
        get
        {
            lock (_sync)
                return _all.Select(w => (w.Pid, w.ExcelPid)).ToList();
        }
    }

    public async Task<WorkerResponse> ConvertAsync(string input, string output, int format, CancellationToken ct)
    {
        var response = await SendAsync(new WorkerRequest("convert", Path.GetFullPath(input), Path.GetFullPath(output), format), ct).ConfigureAwait(false);
        if (!File.Exists(output))
            throw new ExcelException(ExcelErrorKind.Excel, L.T("Excel не створив файл результату."));
        return response;
    }

    public async Task<WorkerResponse> ConvertInspectingAsync(string input, string output, int format, CancellationToken ct)
    {
        var response = await SendAsync(new WorkerRequest("convert", Path.GetFullPath(input), Path.GetFullPath(output), format, Inspect: true), ct).ConfigureAwait(false);
        if (!File.Exists(output))
            throw new ExcelException(ExcelErrorKind.Excel, L.T("Excel не створив файл результату."));
        return response;
    }

    public Task<WorkerResponse> InspectAsync(string input, CancellationToken ct) =>
        SendAsync(new WorkerRequest("inspect", Path.GetFullPath(input)), ct);

    /// <summary>Sends one request to a free worker; for tests also "stall" and "exit".</summary>
    internal async Task<WorkerResponse> SendAsync(WorkerRequest request, CancellationToken ct)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_unavailable is { } reason)
            throw new ExcelException(ExcelErrorKind.Unavailable, reason);
        await _slots.WaitAsync(ct).ConfigureAwait(false);
        WorkerProcess? worker = null;
        Interlocked.Increment(ref _busy);
        try
        {
            worker = TakeIdle() ?? await StartAsync(ct).ConfigureAwait(false);
            WorkerResponse response;
            try
            {
                response = await worker.SendAsync(request, _options.Timeout, ct).ConfigureAwait(false);
            }
            catch (ExcelException ex) when (ex.Kind is ExcelErrorKind.Timeout or ExcelErrorKind.Crashed)
            {
                Interlocked.Increment(ref ex.Kind == ExcelErrorKind.Timeout ? ref _timeouts : ref _crashes);
                Discard(worker);
                worker = null;
                throw;
            }
            catch (OperationCanceledException)
            {
                // Excel cannot be interrupted in the middle of an operation; the worker is not reused.
                Discard(worker);
                worker = null;
                throw;
            }
            Interlocked.Increment(ref _operations);
            if (!response.Ok && response.Kind == ExcelErrorKind.Crashed)
            {
                Interlocked.Increment(ref _crashes);
                Discard(worker);
                worker = null;
            }
            else if (response.Restart == true)
            {
                // Excel asked something (the worker closed the dialog): the next file gets a fresh Excel, none is left on screen.
                Interlocked.Increment(ref _dialogs);
                lock (_sync)
                    _all.Remove(worker);
                await worker.QuitAsync().ConfigureAwait(false);
                worker = null;
            }
            if (!response.Ok)
                throw new ExcelException(response.Kind ?? ExcelErrorKind.Excel, response.Error ?? L.T("Помилка Excel."));
            return response;
        }
        finally
        {
            if (worker is not null)
                await ReturnAsync(worker).ConfigureAwait(false);
            Interlocked.Decrement(ref _busy);
            _slots.Release();
            _recent?.Sweep();
        }
    }

    private WorkerProcess? TakeIdle()
    {
        lock (_sync)
            return _idle.Count > 0 ? _idle.Pop() : null;
    }

    private async Task<WorkerProcess> StartAsync(CancellationToken ct)
    {
        try
        {
            var worker = await WorkerProcess.StartAsync(_options, _registry, ct).ConfigureAwait(false);
            lock (_sync)
                _all.Add(worker);
            return worker;
        }
        catch (ExcelException ex) when (ex.Kind == ExcelErrorKind.Unavailable)
        {
            // No point in starting Excel for every file of the run.
            _unavailable = ex.Message;
            throw;
        }
    }

    private async Task ReturnAsync(WorkerProcess worker)
    {
        if (_disposed || worker.Operations >= _options.RecycleAfter)
        {
            lock (_sync)
                _all.Remove(worker);
            await worker.QuitAsync().ConfigureAwait(false);
            return;
        }
        lock (_sync)
            _idle.Push(worker);
    }

    private void Discard(WorkerProcess worker)
    {
        lock (_sync)
            _all.Remove(worker);
        worker.Kill();
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;
        _disposed = true;
        List<WorkerProcess> all;
        lock (_sync)
        {
            all = [.. _all];
            _all.Clear();
            _idle.Clear();
        }
        await Task.WhenAll(all.Select(w => w.QuitAsync())).ConfigureAwait(false);
        _recent?.Sweep();
        _registry.Dispose();
    }
}

/// <summary>One worker process and the Excel it drives.</summary>
internal sealed class WorkerProcess
{
    private readonly Process _process;
    private readonly ProcessRegistry _registry;
    private readonly StreamWriter _input;
    private readonly StreamReader _output;
    private (int Pid, DateTime Start)? _excel;
    private int _killed;

    private WorkerProcess(Process process, ProcessRegistry registry)
    {
        _process = process;
        Pid = process.Id;
        _registry = registry;
        _input = new StreamWriter(process.StandardInput.BaseStream, new UTF8Encoding(false)) { AutoFlush = true };
        _output = new StreamReader(process.StandardOutput.BaseStream, new UTF8Encoding(false));
    }

    public int Operations { get; private set; }

    public int Pid { get; }

    public int? ExcelPid => _excel?.Pid;

    public static async Task<WorkerProcess> StartAsync(ExcelPoolOptions options, ProcessRegistry registry, CancellationToken ct)
    {
        bool dll = options.WorkerPath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase);
        if (!File.Exists(options.WorkerPath))
            throw new ExcelException(ExcelErrorKind.Unavailable, L.F("Не знайдено обробник Excel: {0}.", options.WorkerPath));
        var start = new ProcessStartInfo(dll ? "dotnet" : options.WorkerPath)
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetDirectoryName(options.WorkerPath)!,
        };
        if (dll)
            start.ArgumentList.Add(options.WorkerPath);
        if (options.Fake)
            start.ArgumentList.Add("--fake");
        // The worker answers in the language of the engine.
        start.Environment["SPO_RL_LANG"] = L.Language;

        Process process;
        try
        {
            process = Process.Start(start) ?? throw new ExcelException(ExcelErrorKind.Unavailable, L.T("Обробник Excel не запустився."));
        }
        catch (System.ComponentModel.Win32Exception ex)
        {
            throw new ExcelException(ExcelErrorKind.Unavailable, L.T("Обробник Excel не запустився: ") + ex.Message, ex);
        }
        registry.Add(process.Id, process.StartTime);
        process.ErrorDataReceived += (_, _) => { };
        process.BeginErrorReadLine();
        var worker = new WorkerProcess(process, registry);
        try
        {
            var deadline = DateTime.UtcNow + options.StartTimeout;
            while (true)
            {
                var left = deadline - DateTime.UtcNow;
                string? line = left > TimeSpan.Zero ? await worker.ReadLineAsync(left, ct).ConfigureAwait(false) : null;
                var ready = line is null ? null : JsonSerializer.Deserialize<WorkerReady>(line, Protocol.Json);
                if (ready is null)
                    throw new ExcelException(ExcelErrorKind.Crashed, L.T("Обробник Excel завершився під час запуску."));
                // The pid comes first, so an Excel stuck while starting is killed together with the worker.
                if (ready.ExcelPid > 0 && worker._excel is null)
                    worker.RememberExcel(ready.ExcelPid);
                if (ready.Error is not null)
                    throw new ExcelException(ready.Kind ?? ExcelErrorKind.Crashed, ready.Error);
                if (ready.Ready)
                    break;
            }
            return worker;
        }
        catch (Exception ex)
        {
            worker.Kill();
            if (ex is ExcelException or OperationCanceledException)
                throw;
            // Not a missing Excel: maybe a one-off failure, the next file tries a new worker.
            throw new ExcelException(ExcelErrorKind.Crashed, L.T("Обробник Excel не запустився: ") + ex.Message, ex);
        }
    }

    private void RememberExcel(int pid)
    {
        try
        {
            using var excel = Process.GetProcessById(pid);
            _excel = (pid, excel.StartTime);
            _registry.Add(pid, excel.StartTime);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            // Excel already exited; the first request reports it.
        }
    }

    public async Task<WorkerResponse> SendAsync(WorkerRequest request, TimeSpan timeout, CancellationToken ct)
    {
        Operations++;
        try
        {
            await _input.WriteLineAsync(JsonSerializer.Serialize(request, Protocol.Json).AsMemory(), ct).ConfigureAwait(false);
        }
        catch (IOException ex)
        {
            throw new ExcelException(ExcelErrorKind.Crashed, L.T("Обробник Excel завершився: ") + ex.Message, ex);
        }
        string? line = await ReadLineAsync(timeout, ct).ConfigureAwait(false);
        if (line is null)
            throw new ExcelException(ExcelErrorKind.Crashed, L.T("Обробник Excel завершився посеред обробки файлу."));
        try
        {
            return JsonSerializer.Deserialize<WorkerResponse>(line, Protocol.Json)
                ?? throw new ExcelException(ExcelErrorKind.Crashed, L.T("Порожня відповідь обробника Excel."));
        }
        catch (JsonException ex)
        {
            // Out of step with the worker: never read another file's answer from it.
            throw new ExcelException(ExcelErrorKind.Crashed, L.T("Незрозуміла відповідь обробника Excel: ") + ex.Message, ex);
        }
    }

    private async Task<string?> ReadLineAsync(TimeSpan timeout, CancellationToken ct)
    {
        var read = _output.ReadLineAsync(CancellationToken.None).AsTask();
        using var delay = CancellationTokenSource.CreateLinkedTokenSource(ct);
        var finished = await Task.WhenAny(read, Task.Delay(timeout, delay.Token)).ConfigureAwait(false);
        delay.Cancel();
        if (finished == read)
        {
            try
            {
                return await read.ConfigureAwait(false);
            }
            catch (IOException)
            {
                return null;
            }
        }
        Kill();
        _ = read.ContinueWith(t => _ = t.Exception, TaskScheduler.Default);
        ct.ThrowIfCancellationRequested();
        throw new ExcelException(ExcelErrorKind.Timeout, L.F("Excel не відповів за {0:0.#} хв; процес завершено примусово.", timeout.TotalMinutes));
    }

    /// <summary>Asks the worker to close Excel; if it does not, kills both.</summary>
    public async Task QuitAsync()
    {
        try
        {
            await _input.WriteLineAsync(JsonSerializer.Serialize(new WorkerRequest("quit"), Protocol.Json)).ConfigureAwait(false);
            using var wait = new CancellationTokenSource(TimeSpan.FromSeconds(20));
            await _process.WaitForExitAsync(wait.Token).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or OperationCanceledException or InvalidOperationException)
        {
        }
        Kill();
    }

    /// <summary>Kills the worker and its Excel if they are still running.</summary>
    public void Kill()
    {
        if (Interlocked.Exchange(ref _killed, 1) == 1)
            return;
        try
        {
            if (!_process.HasExited)
                _process.Kill(entireProcessTree: true);
            _process.WaitForExit(10_000);
        }
        catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
        {
        }
        if (_excel is { } excel)
            ProcessRegistry.KillIfSame(excel.Pid, excel.Start);
        _registry.Remove(_process.Id);
        if (_excel is { } e)
            _registry.Remove(e.Pid);
        _process.Dispose();
    }
}

/// <summary>
/// Remembers the processes a pool started (+pid start / -pid lines). The file stays open while the pool lives;
/// a file nobody holds belongs to a host that crashed, and its processes are killed.
/// </summary>
internal sealed class ProcessRegistry : IDisposable
{
    private readonly FileStream _file;
    private readonly StreamWriter _writer;
    private readonly string _path;
    private readonly object _sync = new();

    public ProcessRegistry(string dir)
    {
        Directory.CreateDirectory(dir);
        KillOrphans(dir);
        _path = Path.Combine(dir, $"excel-{Guid.NewGuid():N}.pids");
        _file = new FileStream(_path, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
        _writer = new StreamWriter(_file) { AutoFlush = true };
    }

    public void Add(int pid, DateTime start)
    {
        lock (_sync)
            _writer.WriteLine($"+{pid} {start.ToUniversalTime().Ticks}");
    }

    public void Remove(int pid)
    {
        lock (_sync)
        {
            if (_file.CanWrite)
                _writer.WriteLine($"-{pid}");
        }
    }

    /// <summary>Kills what pools of hosts that are no longer running left behind.</summary>
    public static int KillOrphans(string dir)
    {
        int killed = 0;
        foreach (string path in Directory.EnumerateFiles(dir, "excel-*.pids"))
        {
            string[] lines;
            try
            {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                using var reader = new StreamReader(stream);
                lines = reader.ReadToEnd().Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            }
            catch (IOException)
            {
                continue; // a live pool holds it
            }
            var alive = new Dictionary<int, long>();
            foreach (string line in lines)
            {
                string[] parts = line[1..].Split(' ');
                if (!int.TryParse(parts[0], out int pid))
                    continue;
                if (line[0] == '+' && parts.Length > 1 && long.TryParse(parts[1], out long ticks))
                    alive[pid] = ticks;
                else if (line[0] == '-')
                    alive.Remove(pid);
            }
            foreach (var (pid, ticks) in alive)
            {
                if (KillIfSame(pid, new DateTime(ticks, DateTimeKind.Utc)))
                    killed++;
            }
            try
            {
                File.Delete(path);
            }
            catch (IOException)
            {
            }
        }
        return killed;
    }

    /// <summary>Kills the process only if it is the same one (pid and start time), so a reused pid is never touched.</summary>
    public static bool KillIfSame(int pid, DateTime start)
    {
        try
        {
            using var process = Process.GetProcessById(pid);
            if (process.StartTime.ToUniversalTime().Ticks != start.ToUniversalTime().Ticks)
                return false;
            process.Kill(entireProcessTree: true);
            process.WaitForExit(10_000);
            return true;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            return false;
        }
    }

    public void Dispose()
    {
        lock (_sync)
        {
            _writer.Dispose();
            _file.Dispose();
        }
        try
        {
            File.Delete(_path);
        }
        catch (IOException)
        {
        }
    }
}
