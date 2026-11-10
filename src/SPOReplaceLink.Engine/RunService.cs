using Microsoft.Extensions.Logging;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Rules;
using SPOReplaceLink.Spo;
using SPOReplaceLink.State;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Engine;

public static class SettingKinds
{
    public const string Profile = "profile";
    public const string Connection = "connection";
    public const string Template = "template";
}

public sealed class RunValidationException(IReadOnlyList<ValidationIssue> issues)
    : Exception(L.T("Профіль містить помилки: ") + string.Join("; ", issues.Where(i => i.Severity == IssueSeverity.Error)))
{
    public IReadOnlyList<ValidationIssue> Issues { get; } = issues;
}

/// <summary>
/// The engine's front door for the host API and the GUI: stored settings, creation of runs from profiles,
/// retry runs (RTY-4), and control of the one active run (start, pause, continue, stop).
/// </summary>
public sealed class RunService : IAsyncDisposable
{
    private readonly StateStore _store;
    private readonly ISpoAppFactory _apps;
    private readonly ILoggerFactory _loggers;
    private readonly EngineOptions _options;
    private readonly object _sync = new();
    private ActiveRun? _active;

    public RunService(StateStore store, ISpoAppFactory apps, ILoggerFactory loggers, EngineOptions options)
    {
        _store = store;
        _apps = apps;
        _loggers = loggers;
        _options = options;
    }

    public StateStore Store => _store;

    private DateTimeOffset Now => _options.Time.GetUtcNow();

    public sealed class ActiveRun(long runId, RunEngine engine, CancellationTokenSource stop, CancellationTokenSource abort)
    {
        public long RunId { get; } = runId;
        public RunEngine Engine { get; } = engine;
        public CancellationTokenSource Stop { get; } = stop;
        public CancellationTokenSource Abort { get; } = abort;
        public Task<string> Completion { get; set; } = Task.FromResult(RunStatuses.Created);
    }

    public ActiveRun? Active
    {
        get
        {
            lock (_sync)
                return _active;
        }
    }

    // ------------------------------------------------------------------ settings

    public IReadOnlyDictionary<string, Template> Templates() =>
        _store.ListSettings(SettingKinds.Template).ToDictionary(t => t.Name, t => Template.Parse(t.Json), StringComparer.OrdinalIgnoreCase);

    public IReadOnlyList<Connection> Connections() =>
        _store.ListSettings(SettingKinds.Connection).Select(c => Connection.Parse(c.Json)).ToList();

    public Profile GetProfile(string name) =>
        _store.GetSetting(SettingKinds.Profile, name) is { } json ? Profile.Parse(json) : throw new KeyNotFoundException(L.F("Профіль «{0}» не знайдено.", name));

    /// <summary>Validates a profile as a run would: templates resolved, connections present (GUI-27).</summary>
    public ProfileValidation ValidateProfile(Profile profile)
    {
        Profile resolved;
        try
        {
            resolved = profile.ResolveTemplates(Templates());
        }
        catch (ProfileException ex)
        {
            return new ProfileValidation(null, [new ValidationIssue(IssueSeverity.Error, ex.Message)]);
        }
        var used = Connections().Where(c => profile.Connections.Contains(c.Name, StringComparer.OrdinalIgnoreCase)).ToList();
        return resolved.ValidateForRun(used);
    }

    public sealed record ConnectionCheck(string What, bool Ok, string Message);

    /// <summary>GUI-26: gets a token and, with a profile, checks access to each of its targets.</summary>
    public async Task<IReadOnlyList<ConnectionCheck>> TestConnectionAsync(string name, string? profileName, CancellationToken ct)
    {
        var connection = Connections().FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException(L.F("Підключення «{0}» не знайдено.", name));
        var checks = new List<ConnectionCheck>();
        var invalid = connection.Validate().Where(i => i.Severity == IssueSeverity.Error).ToList();
        if (invalid.Count > 0)
            return [new ConnectionCheck(L.T("налаштування"), false, string.Join("; ", invalid.Select(i => i.Message)))];

        SpoApp app;
        try
        {
            var tenant = new Spo.Throttling.RuBudget("tenant", connection.Limits().TenantPer5Min * 0.5, TimeSpan.FromMinutes(5), 0, _options.Time);
            app = _apps.Create(connection, new SpoAppSettings(4, 3, TimeSpan.FromSeconds(60), TimeSpan.FromMinutes(15), 0.5), tenant);
        }
        catch (SpoException ex)
        {
            return [new ConnectionCheck(L.T("сертифікат"), false, ex.Message)];
        }

        var targets = profileName is null ? [] : GetProfile(profileName).Targets;
        try
        {
            await app.Client.CheckTokenAsync(ct).ConfigureAwait(false);
            checks.Add(new ConnectionCheck(L.T("токен"), true, L.T("Токен отримано.")));
        }
        catch (SpoException ex)
        {
            checks.Add(new ConnectionCheck(L.T("токен"), false, ex.Message));
            return checks;
        }
        foreach (var target in targets)
        {
            try
            {
                var (drive, _, root) = await TargetResolver.ResolveAsync(app.Client, target, ct).ConfigureAwait(false);
                checks.Add(new ConnectionCheck(target.ToString(), true, L.F("Доступ є: бібліотека {0}{1}.", drive.Name, root.Length > 0 ? L.T(", папка ") + root : string.Empty)));
            }
            catch (SpoException ex)
            {
                checks.Add(new ConnectionCheck(target.ToString(), false, ex.Message));
            }
        }
        return checks;
    }

    // ------------------------------------------------------------------ runs

    /// <summary>New run of a stored profile (or of a profile given inline) in the given mode.</summary>
    public long CreateRun(string mode, string? profileName, Profile? inline = null)
    {
        if (!RunModes.All.Contains(mode))
            throw new ArgumentException(L.F("Невідомий режим «{0}» ({1}).", mode, string.Join(", ", RunModes.All)));
        var profile = inline ?? GetProfile(profileName ?? throw new ArgumentException(L.T("Не задано профіль.")));
        var config = Snapshot(profile);
        return _store.CreateRun(mode, null, profileName ?? profile.Name, config.ToJson(), enumerate: true, config.Profile.Targets.Count, Now);
    }

    /// <summary>
    /// RTY-4 and verify after apply: a new run over the files of <paramref name="parentId"/> picked by the filter;
    /// the library is not listed again, every file is downloaded fresh (RTY-5).
    /// </summary>
    public long CreateChildRun(long parentId, FileFilter filter, string? mode = null, bool currentProfile = false)
    {
        var parent = _store.GetRun(parentId) ?? throw new KeyNotFoundException(L.F("Прогін {0} не знайдено.", parentId));
        mode ??= parent.Mode;
        if (!RunModes.All.Contains(mode))
            throw new ArgumentException(L.F("Невідомий режим «{0}».", mode));
        string snapshot = parent.ConfigSnapshot;
        if (currentProfile && parent.ProfileName is { } name)
        {
            var fresh = Snapshot(GetProfile(name));
            // Files keep the target index of the parent run: the targets must be the same, only rules and settings may differ.
            var old = RunConfig.Parse(parent.ConfigSnapshot).Profile.Targets;
            bool same = old.Count == fresh.Profile.Targets.Count && old.Zip(fresh.Profile.Targets).All(p =>
                string.Equals(p.First.SiteUrl?.TrimEnd('/'), p.Second.SiteUrl?.TrimEnd('/'), StringComparison.OrdinalIgnoreCase)
                && string.Equals(p.First.Library, p.Second.Library, StringComparison.OrdinalIgnoreCase)
                && string.Equals((p.First.Folder ?? string.Empty).Trim('/'), (p.Second.Folder ?? string.Empty).Trim('/'), StringComparison.OrdinalIgnoreCase));
            if (!same)
                throw new InvalidOperationException(L.T("Цілі профілю змінились після прогону; для нових цілей запустіть новий прогін, а не повтор."));
            snapshot = fresh.ToJson();
        }
        long id = _store.CreateRun(mode, parentId, parent.ProfileName, snapshot, enumerate: false, targets: 0, Now);
        int copied = _store.CopyFiles(parentId, id, filter, Now);
        if (copied == 0)
        {
            _store.SetRunStatus(id, RunStatuses.Completed, Now, L.T("Жоден файл не відповідає фільтру."));
        }
        return id;
    }

    private RunConfig Snapshot(Profile profile)
    {
        var resolved = profile.ResolveTemplates(Templates());
        var connections = Connections().Where(c => profile.Connections.Contains(c.Name, StringComparer.OrdinalIgnoreCase)).ToList();
        var validation = resolved.ValidateForRun(connections);
        if (!validation.IsValid)
            throw new RunValidationException(validation.Issues);
        return new RunConfig { Profile = resolved, Connections = connections };
    }

    /// <summary>Starts or continues a run in the background. Only one run is active at a time: they share app budgets.</summary>
    public ActiveRun Start(long runId)
    {
        var run = _store.GetRun(runId) ?? throw new KeyNotFoundException(L.F("Прогін {0} не знайдено.", runId));
        if (!RunStatuses.CanContinue(run.Status))
            throw new InvalidOperationException(L.F("Прогін {0} уже завершено ({1}).", runId, run.Status));
        lock (_sync)
        {
            if (_active is { } current && !current.Completion.IsCompleted)
            {
                if (current.RunId == runId)
                    return current;
                throw new InvalidOperationException(L.F("Уже виконується прогін {0}. Зупиніть його або дочекайтесь завершення.", current.RunId));
            }
            var engine = new RunEngine(_store, _apps, _loggers.CreateLogger($"Run.{runId}"), _options);
            var active = new ActiveRun(runId, engine, new CancellationTokenSource(), new CancellationTokenSource());
            active.Completion = Task.Run(async () =>
            {
                try
                {
                    return await engine.RunAsync(runId, active.Stop.Token, active.Abort.Token).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    _loggers.CreateLogger("RunService").LogError(ex, L.T("Прогін {Run} завершився аварійно"), runId);
                    _store.SetRunStatus(runId, RunStatuses.Aborted, Now, L.T("Внутрішня помилка: ") + ex.Message);
                    return RunStatuses.Aborted;
                }
            });
            _active = active;
            return active;
        }
    }

    /// <summary>
    /// Deletes a run that is not running: its files, results and listing. The list of runs, the estimate of a new run
    /// and the reports no longer show it. What it wrote to SharePoint stays there; only the record of it goes.
    /// </summary>
    public void DeleteRun(long runId)
    {
        RunRow run;
        lock (_sync)
        {
            run = _store.GetRun(runId) ?? throw new KeyNotFoundException(L.F("Прогін {0} не знайдено.", runId));
            if (_active is { } current && current.RunId == runId && !current.Completion.IsCompleted)
                throw new InvalidOperationException(L.F("Прогін {0} виконується. Зупиніть його й дочекайтесь зупинки, потім видаліть.", runId));
            // Marked running or paused but not active: the host was killed, and the run continues on the next start (STA-3).
            if (run.Status is RunStatuses.Running or RunStatuses.Paused)
                throw new InvalidOperationException(L.F("Прогін {0} позначено як «{1}»: він продовжиться після перезапуску. Спершу зупиніть його.", runId, run.Status));
            _store.DeleteRun(runId);
        }
        _loggers.CreateLogger("RunService").LogWarning(L.T("Прогін {Run} ({Mode}, профіль {Profile}, {Status}) видалено"), runId, run.Mode, run.ProfileName, run.Status);
    }

    /// <summary>
    /// STA-3: after the host was killed, a run still marked running (or paused) is continued from where it stopped.
    /// </summary>
    public long? ContinueInterrupted()
    {
        var run = _store.ListRuns(200).Where(r => r.Status is RunStatuses.Running or RunStatuses.Paused).OrderBy(r => r.Id).LastOrDefault();
        if (run is null)
            return null;
        _loggers.CreateLogger("RunService").LogWarning(L.T("Продовження перерваного прогону {Run}"), run.Id);
        Start(run.Id);
        return run.Id;
    }

    public void Pause(long runId)
    {
        var active = Require(runId);
        active.Engine.Pause();
        _store.SetRunStatusIf(runId, RunStatuses.Paused, Now, RunStatuses.Running);
    }

    public void Resume(long runId)
    {
        var active = Require(runId);
        active.Engine.Resume();
        _store.SetRunStatusIf(runId, RunStatuses.Running, Now, RunStatuses.Paused);
    }

    /// <summary>STA-4: graceful stop; the run can be continued later with <see cref="Start"/>.</summary>
    public void Stop(long runId) => Require(runId).Stop.Cancel();

    private ActiveRun Require(long runId)
    {
        var active = Active;
        if (active is null || active.RunId != runId || active.Completion.IsCompleted)
            throw new InvalidOperationException(L.F("Прогін {0} зараз не виконується.", runId));
        return active;
    }

    /// <summary>Host shutdown: graceful stop, then a hard stop if files do not finish in time.</summary>
    public async ValueTask DisposeAsync()
    {
        var active = Active;
        if (active is null || active.Completion.IsCompleted)
            return;
        active.Stop.Cancel();
        if (await Task.WhenAny(active.Completion, Task.Delay(TimeSpan.FromSeconds(30))).ConfigureAwait(false) != active.Completion)
            active.Abort.Cancel();
        try
        {
            await active.Completion.ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
        }
    }
}
