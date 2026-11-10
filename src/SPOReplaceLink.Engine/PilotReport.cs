using System.Globalization;
using System.Text;
using System.Text.Json;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.State;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Engine;

/// <summary>
/// Stage 5 (pilot): what a run measured about SharePoint — RU per file, throttling, speed and a forecast — and the
/// questions only a real tenant answers (version numbers in $select, files changed by SharePoint on save).
/// <see cref="Sample"/> picks changed files for the manual check. Both read the files in one pass, so they work on
/// runs of hundreds of thousands of files.
/// </summary>
public static class PilotReport
{
    /// <summary>THR-4 targets: RU per file without changes and per changed file.</summary>
    private const double RuUnchanged = 2;
    private const double RuChanged = 4;

    /// <summary>
    /// Totals over all sessions of a run (each start of a run counts from zero) and the time files were actually being
    /// finished: minutes spent waiting for a time window, a daily budget or between a stop and a continue are left out.
    /// </summary>
    public sealed record LoadTotals(long Ru, long Requests, long Throttled429, long Throttled503, long ResizedBySpo,
        long ExcelOperations, long ExcelTimeouts, long ExcelCrashes, IReadOnlyDictionary<string, long> RuByApp, TimeSpan ActiveTime);

    public static LoadTotals Load(IEnumerable<string> snapshots)
    {
        // Counters only grow within a session, so a session's total is its largest value.
        var sessions = new List<Dictionary<string, long>>();
        Dictionary<string, long>? current = null;
        DateTimeOffset? session = null;
        MetricsSnapshot? previous = null;
        var active = TimeSpan.Zero;
        foreach (string json in snapshots)
        {
            MetricsSnapshot? s;
            try
            {
                s = JsonSerializer.Deserialize<MetricsSnapshot>(json, Profile.JsonOptions);
            }
            catch (JsonException)
            {
                continue;
            }
            if (s is null)
                continue;
            // A new session: another start time, or (snapshots without it) a counter that went down.
            bool fresh = current is null || previous is null
                || (s.Started is not null ? s.Started != session : s.Finished < previous.Finished || s.Apps.Sum(a => a.RuSpent) < previous.Apps.Sum(a => a.RuSpent));
            if (fresh)
            {
                current = new Dictionary<string, long>(StringComparer.Ordinal);
                sessions.Add(current);
                session = s.Started;
                if (s.Started is { } started && s.Finished > s.FinishedBefore) // the session finished files of its own
                    active += s.At - started;
            }
            else if (s.Finished > previous!.Finished)
            {
                active += s.At - previous.At;
            }
            previous = s;

            void Max(string key, long value) => current![key] = Math.Max(current.GetValueOrDefault(key), value);
            foreach (var a in s.Apps)
            {
                Max("ru:" + a.Name, a.RuSpent);
                Max("req:" + a.Name, a.Requests);
                Max("429:" + a.Name, a.Throttled429);
                Max("503:" + a.Name, a.Throttled503);
            }
            Max("resized", s.ResizedBySpo);
            if (s.Excel is { } x)
            {
                Max("xl:ops", x.Operations);
                Max("xl:timeouts", x.Timeouts);
                Max("xl:crashes", x.Crashes);
            }
        }
        long Sum(string prefix) => sessions.SelectMany(d => d).Where(x => x.Key.StartsWith(prefix, StringComparison.Ordinal)).Sum(x => x.Value);
        var byApp = sessions.SelectMany(d => d).Where(x => x.Key.StartsWith("ru:", StringComparison.Ordinal))
            .GroupBy(x => x.Key[3..]).ToDictionary(g => g.Key, g => g.Sum(x => x.Value));
        return new LoadTotals(Sum("ru:"), Sum("req:"), Sum("429:"), Sum("503:"), Sum("resized"), Sum("xl:ops"), Sum("xl:timeouts"), Sum("xl:crashes"),
            byApp, active);
    }

    private sealed class FormatStat
    {
        public long Files;
        public long Bytes;
        public long Changed;
    }

    public static string Text(StateStore store, long runId)
    {
        var run = store.GetRun(runId) ?? throw new ArgumentException(L.F("Прогін {0} не знайдено.", runId));
        var config = RunConfig.Parse(run.ConfigSnapshot);
        var profile = config.Profile;
        bool apply = run.Mode == RunModes.Apply;

        // One pass over the files.
        long total = 0, touched = 0, changed = 0, written = 0, withVersion = 0, unchanged = 0, replacements = 0, withResiduals = 0, needsExcel = 0;
        var formats = new Dictionary<string, FormatStat>(StringComparer.Ordinal);
        var attention = new Dictionary<(string Code, string? Error), (long Count, FileRow First)>();
        foreach (var f in store.AllFiles(runId))
        {
            string code = f.StatusCode ?? f.Status;
            bool isChanged = apply ? f.Status == FileStatuses.Done : f.Replacements > 0;
            total++;
            if (f.Attempts > 0)
                touched++;
            if (isChanged)
                changed++;
            if (f.Status == FileStatuses.Done)
            {
                written++;
                if (!string.IsNullOrEmpty(f.VersionBefore))
                    withVersion++;
            }
            if (code.StartsWith(FileStatuses.Unchanged, StringComparison.Ordinal) || code.StartsWith(FileStatuses.Verified, StringComparison.Ordinal))
                unchanged++;
            replacements += f.Replacements;
            if (f.Residuals > 0)
                withResiduals++;
            if (code.Contains("needs_excel", StringComparison.Ordinal))
                needsExcel++;
            string ext = Path.GetExtension(f.Name).TrimStart('.').ToLowerInvariant();
            if (!formats.TryGetValue(ext, out var fs))
                formats[ext] = fs = new FormatStat();
            fs.Files++;
            fs.Bytes += f.Size;
            if (isChanged)
                fs.Changed++;
            if (ReportExporter.NeedsAttention(f))
            {
                var key = (code, f.Error);
                attention[key] = attention.TryGetValue(key, out var a) ? (a.Count + 1, a.First) : (1, f);
            }
        }
        var load = Load(store.GetMetrics(runId).Select(m => m.Json));
        var attempts = store.CountAttemptsByClass(runId);
        var duration = (run.FinishedAt ?? DateTimeOffset.UtcNow) - (run.StartedAt ?? run.CreatedAt);

        var sb = new StringBuilder();
        void Line(string text = "") => sb.AppendLine(text);
        static string N(double v, string format = "0") => v.ToString(format, CultureInfo.InvariantCulture);

        Line(L.F("Звіт пілоту: прогін {0} ({1}), профіль {2}", run.Id, run.Mode, run.ProfileName));
        Line(L.F("Статус: {0}{1}", run.Status, run.StopReason is null ? string.Empty : " — " + run.StopReason));
        Line(L.F("Цілі: {0}", string.Join("; ", profile.Targets)));
        Line(L.F("Від першого старту до завершення: {0:d\\.hh\\:mm\\:ss}; активна обробка: {1:d\\.hh\\:mm\\:ss}", duration, load.ActiveTime));
        Line();

        Line(L.T("1. Файли"));
        Line(L.F("  у прогоні: {0}; оброблено (завантажено хоча б раз): {1}", total, touched));
        Line(L.F("  {0}: {1}; без змін: {2}", apply ? L.T("записано в SharePoint") : L.T("буде змінено"), changed, unchanged));
        Line(L.F("  замін посилань: {0}; файлів із залишками: {1}", replacements, withResiduals));
        foreach (var (ext, fs) in formats.OrderByDescending(x => x.Value.Files))
            Line(L.F("  .{0}: {1} файлів, {2} МБ, змінено {3}", ext, fs.Files, N(fs.Bytes / 1024.0 / 1024.0, "0.0"), fs.Changed));
        Line();

        Line(L.T("2. Навантаження на SharePoint"));
        Line(L.F("  RU витрачено: {0}{1}; запитів: {2}", load.Ru, load.RuByApp.Count > 1 ? " (" + string.Join(", ", load.RuByApp.Select(x => $"{x.Key}: {x.Value}")) + ")" : string.Empty, load.Requests));
        Line(L.F("  відповідей 429: {0}; 503: {1}", load.Throttled429, load.Throttled503));
        double ruPerFile = touched > 0 ? (double)load.Ru / touched : 0;
        if (touched > 0)
        {
            double expected = apply ? (changed * RuChanged + (touched - changed) * RuUnchanged) / touched : RuUnchanged;
            Line(L.F("  RU на файл: {0} (разом із переліком бібліотеки); орієнтир плану: {1}", N(ruPerFile, "0.00"), N(expected, "0.0")));
            if (ruPerFile > expected * 1.5)
                Line(L.T("  ! Вартість файлу помітно вища за план: перевірте частку повторів (attempts) і відповіді 412/423 у журналі."));
        }
        if (load.Throttled429 > 0)
            Line(L.T("  ! Були 429 — знизьте бюджет RU (performance.budgetPercent і бюджети вікон) або перенесіть прогін у неробочий час."));
        if (load.Throttled503 > 0)
            Line(L.T("  ! Були 503: SharePoint перевантажений. Якщо їх багато, працюйте в неробочий час."));
        Line();

        Line(L.T("3. Швидкість і прогноз"));
        double perFileSeconds = touched > 0 && load.ActiveTime > TimeSpan.Zero ? load.ActiveTime.TotalSeconds / touched : 0;
        double measured = perFileSeconds > 0 ? 3600 / perFileSeconds : 0;
        if (measured > 0)
            Line(L.F("  виміряно за час активної обробки: {0} файлів/год ({1} с на файл)", N(measured), N(perFileSeconds, "0.00")));
        else
            Line(L.T("  швидкість не виміряно: замало знімків метрик (вони пишуться раз на хвилину)."));
        double ceiling = Ceiling(config, ruPerFile, out string budgetText);
        if (ceiling > 0)
            Line(L.F("  межа бюджету RU ({0}): ~{1} файлів/год", budgetText, N(ceiling)));
        double rate = measured > 0 && ceiling > 0 ? Math.Min(measured, ceiling) : Math.Max(measured, ceiling);
        if (rate > 0)
        {
            Line(L.F("  орієнтовно на 10 000 файлів: {0}", Hours(10_000 / rate)) + (ceiling > 0 && measured > ceiling ? L.T(" (обмежує бюджет RU)") : string.Empty));
            if (ProfileInsights.For(store, profile) is { PassingFilters: > 0 } estimate)
                Line(L.F("  на профіль «{0}» ({1} файлів за переліком прогону {2}): {3}", profile.Name, estimate.PassingFilters, estimate.RunId, Hours(estimate.PassingFilters / rate)));
            Line(L.T("  для інших бібліотек: кількість їхніх файлів × с на файл; швидкість залежить від розміру файлів і частки змінених."));
        }
        Line();

        Line(L.T("4. Як SharePoint поводиться з записаними файлами"));
        if (!apply)
        {
            Line(L.T("  прогін нічого не записував — ці перевірки дає прогін apply."));
        }
        else
        {
            Line(L.F("  номер версії до запису відомий: {0} з {1}", withVersion, written));
            if (written > 0 && withVersion == 0)
                Line(L.T("  ! SharePoint не віддає номер версії (поля publication / listItem у $select): відкат за версією і «копія версії до зміни» недоступні. Передайте звіт розробнику."));
            Line(L.F("  SharePoint змінив розмір файлу під час збереження (за відповіддю на запис): {0} з {1}", load.ResizedBySpo, written));
            if (load.ResizedBySpo > 0)
                Line(L.T("  ! SharePoint дописує у файли властивості бібліотеки. Подвійного запису не буде, але файл, відповідь на запис якого загубилась, отримає unchanged замість done."));
            Line(L.F("  спроб із конфліктом (412): {0}; із блокуванням (423, check-out): {1}", attempts.GetValueOrDefault("conflict"), attempts.GetValueOrDefault("locked")));
        }
        Line();

        if (load.ExcelOperations > 0 || needsExcel > 0)
        {
            Line(L.T("5. Excel (xls, xlsb)"));
            Line(L.F("  операцій: {0}; зависань: {1}; падінь: {2}; файлів needs_excel: {3}", load.ExcelOperations, load.ExcelTimeouts, load.ExcelCrashes, needsExcel));
            Line();
        }

        Line(L.T("6. Помилки і що з ними робити"));
        if (attention.Count == 0)
            Line(L.T("  немає файлів, що потребують уваги."));
        foreach (var (key, (count, first)) in attention.OrderByDescending(x => x.Value.Count).Take(15))
            Line($"  {count,6} × {key.Code}{(string.IsNullOrEmpty(key.Error) ? string.Empty : ": " + Trim(key.Error))} → {ReportExporter.Recommend(first)}");
        Line();

        Line(L.T("7. Перевірити вручну"));
        Line(L.T("  - звіт prefixes: префікси з no_rule і unresolved — кандидати в нові правила;"));
        Line(L.T("  - звіт sample: відкрити кожен файл вибірки, переконатися, що посилання ведуть куди треба (стовпці checked, comment);"));
        if (apply)
            Line(L.T("  - попросити власників ключових файлів відкрити їх у Excel: зв'язки оновлюються, гіперпосилання відкриваються, немає попереджень про відновлення;"));
        Line(L.T("  - verify після apply: 0 файлів changes_pending і 0 залишків."));
        return sb.ToString();
    }

    /// <summary>
    /// Files per hour the RU budget allows: the largest budget the profile works at (window budgets included), limited by
    /// the per-minute and the 24-hour limit of each app and, for a pool, by the tenant ceiling.
    /// </summary>
    private static double Ceiling(RunConfig config, double ruPerFile, out string text)
    {
        text = string.Empty;
        var profile = config.Profile;
        var apps = profile.Connections
            .Select(n => config.Connections.FirstOrDefault(c => string.Equals(c.Name, n, StringComparison.OrdinalIgnoreCase)))
            .OfType<Connection>()
            .ToList();
        if (apps.Count == 0 || ruPerFile <= 0)
            return 0;
        var perf = profile.Performance;
        int percent = perf.Windows.Select(w => w.BudgetPercent)
            .Concat(perf.OnlyInWindows && perf.Windows.Count > 0 ? [] : [perf.BudgetPercent])
            .Max();
        double perMinute = apps.Sum(a => Math.Min(a.Limits().PerMinute, a.Limits().PerDay / 1440.0)) * percent / 100.0;
        if (apps.Count > 1)
            perMinute = Math.Min(perMinute, apps.Max(a => a.Limits().TenantPer5Min) * profile.Pool.TenantCeilingPercent / 100.0 / 5);
        text = L.F("{0}% ліміту з урахуванням добового, {1} RU/хв", percent, N0(perMinute));
        return perMinute * 60 / ruPerFile;
    }

    /// <summary>
    /// Changed files for the manual check, in a fixed pseudo-random order (the same run gives the same sample), with
    /// a link to the file and the first replacements.
    /// </summary>
    public static IEnumerable<IReadOnlyList<string?>> Sample(StateStore store, long runId, int size = 30)
    {
        var targets = store.GetTargets(runId).ToDictionary(t => t.Index);
        // The smallest keys win; only `size` rows are kept however large the run is.
        var top = new PriorityQueue<FileRow, ulong>(Comparer<ulong>.Create((a, b) => b.CompareTo(a)));
        foreach (var f in store.AllFiles(runId))
        {
            if (f.Replacements == 0 || (f.StatusCode ?? f.Status).StartsWith("failed:", StringComparison.Ordinal))
                continue;
            ulong key = Mix(f.Id, runId);
            if (top.Count < size)
                top.Enqueue(f, key);
            else if (size > 0 && top.TryPeek(out _, out ulong worst) && key < worst)
                top.EnqueueDequeue(f, key);
        }
        var picked = new List<(FileRow Row, ulong Key)>();
        while (top.TryDequeue(out var row, out ulong key))
            picked.Add((row, key));

        yield return ["id", "path", "status", "replacements", "residuals", "version_before", "url", "examples", "checked", "comment"];
        foreach (var (f, _) in picked.OrderBy(p => p.Key))
        {
            string? url = targets.TryGetValue(f.TargetIndex, out var t) && t.DriveUrl is { } drive && f.Path is { } path
                ? drive.TrimEnd('/') + "/" + string.Join('/', path.Split('/').Select(Uri.EscapeDataString))
                : null;
            var examples = store.GetLinks(f.Id)
                .Where(l => l.Action == LinkAction.Replaced && l.Kind != LinkKind.WorkbookPath)
                .Take(3)
                .Select(l => $"{l.Kind}: {l.OldValue} → {l.NewValue}");
            yield return [f.Id.ToString(CultureInfo.InvariantCulture), f.Path ?? f.Name, f.StatusCode ?? f.Status,
                f.Replacements.ToString(CultureInfo.InvariantCulture), f.Residuals.ToString(CultureInfo.InvariantCulture),
                f.VersionBefore, url, string.Join("\n", examples), string.Empty, string.Empty];
        }
    }

    private static ulong Mix(long id, long seed)
    {
        ulong x = (ulong)id * 0x9E3779B97F4A7C15UL ^ (ulong)seed;
        x ^= x >> 31;
        x *= 0xBF58476D1CE4E5B9UL;
        return x ^ (x >> 29);
    }

    private static string N0(double v) => v.ToString("0", CultureInfo.InvariantCulture);

    private static string Hours(double hours) => hours < 1
        ? L.F("~{0} хв", Math.Max(1, Math.Round(hours * 60)).ToString(CultureInfo.InvariantCulture))
        : L.F("~{0} год", hours.ToString("0.0", CultureInfo.InvariantCulture));

    private static string Trim(string? s) => s is null ? string.Empty : s.Length <= 160 ? s : s[..160] + "…";
}
