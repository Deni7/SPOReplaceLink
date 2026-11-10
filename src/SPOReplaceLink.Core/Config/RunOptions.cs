using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using SPOReplaceLink.Core.Rules;
using SPOReplaceLink.Core.Urls;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Core.Config;

/// <summary>CFG-2: a library or a folder in it.</summary>
public sealed class Target
{
    /// <summary>https://tenant.sharepoint.com/sites/x</summary>
    public string? SiteUrl { get; set; }

    /// <summary>Library name as shown in SharePoint ("Documents"), its URL name ("Shared Documents") or full URL.</summary>
    public string? Library { get; set; }

    /// <summary>Folder inside the library, e.g. "Fin/2024". Empty — the whole library.</summary>
    public string? Folder { get; set; }

    /// <summary>
    /// Where this library/folder lived before migration, e.g. https://tenant.sharepoint.com/Document — used for relative
    /// links (RUL-6). If empty, the old location is derived by reversing the rules, when that is unambiguous.
    /// </summary>
    public string? OldLocation { get; set; }

    public override string ToString() =>
        $"{SiteUrl?.TrimEnd('/')}/{Library}{(string.IsNullOrEmpty(Folder) ? string.Empty : "/" + Folder.Trim('/'))}";

    internal IEnumerable<ValidationIssue> Validate(int index)
    {
        if (string.IsNullOrWhiteSpace(SiteUrl) || LinkClassifier.Classify(SiteUrl) != LinkForm.Http)
            yield return new(IssueSeverity.Error, L.F("targets[{0}].siteUrl: потрібна адреса сайту https://…", index));
        if (string.IsNullOrWhiteSpace(Library))
            yield return new(IssueSeverity.Error, L.F("targets[{0}].library: не задано бібліотеку.", index));
        if (OldLocation is { Length: > 0 } old && !LinkClassifier.IsAbsolute(LinkClassifier.Classify(old)))
            yield return new(IssueSeverity.Error, L.F("targets[{0}].oldLocation: потрібна повна адреса.", index));
    }
}

/// <summary>THR-9…THR-11.</summary>
public sealed class PoolOptions
{
    /// <summary>Several connections in one run. Off by default: Microsoft advises against several apps for one job.</summary>
    public bool Enabled { get; set; }

    /// <summary>THR-10: share of the tenant budget (RU per 5 min) the whole run may use.</summary>
    public int TenantCeilingPercent { get; set; } = 50;

    internal IEnumerable<ValidationIssue> Validate()
    {
        if (TenantCeilingPercent is < 1 or > 100)
            yield return new(IssueSeverity.Error, L.T("pool.tenantCeilingPercent: від 1 до 100."));
    }
}

/// <summary>CFG-5, THR-1, THR-5.</summary>
public sealed class PerformanceOptions
{
    /// <summary>Files downloaded at the same time.</summary>
    public int Downloads { get; set; } = 8;

    /// <summary>Files processed (CPU) at the same time.</summary>
    public int Processors { get; set; } = Math.Max(2, Environment.ProcessorCount / 2);

    /// <summary>Files uploaded at the same time.</summary>
    public int Uploads { get; set; } = 4;

    public int RequestTimeoutSeconds { get; set; } = 120;

    /// <summary>Whole cycle of one file; exceeded — transient failure.</summary>
    public int FileTimeoutMinutes { get; set; } = 30;

    /// <summary>Larger files are uploaded through an upload session in chunks (UPL-3).</summary>
    public int SimpleUploadMaxMb { get; set; } = 60;

    /// <summary>XLS-4: Excel processes for xls and xlsb residuals. 0 — Excel is not used, such files get needs_excel.</summary>
    public int ExcelProcesses { get; set; } = 2;

    /// <summary>XLS-4: one Excel operation (open and save) longer than this — Excel is considered hung and is killed.</summary>
    public int ExcelTimeoutMinutes { get; set; } = 5;

    /// <summary>THR-1: share of the app RU limit to work at.</summary>
    public int BudgetPercent { get; set; } = 80;

    /// <summary>THR-5: time windows with their own budget. Outside all windows — <see cref="BudgetPercent"/>, or a pause with <see cref="OnlyInWindows"/>.</summary>
    public List<ScheduleWindow> Windows { get; set; } = [];

    public bool OnlyInWindows { get; set; }

    /// <summary>Time zone of the windows (Windows or IANA id). Empty — the host's local time.</summary>
    public string? TimeZone { get; set; }

    public TimeZoneInfo ResolveTimeZone() =>
        string.IsNullOrWhiteSpace(TimeZone) ? TimeZoneInfo.Local : TimeZoneInfo.FindSystemTimeZoneById(TimeZone);

    /// <summary>Budget percent for this moment, or null when the run must wait (outside windows with OnlyInWindows).</summary>
    public int? BudgetAt(DateTimeOffset now)
    {
        var local = TimeZoneInfo.ConvertTime(now, ResolveTimeZone());
        var window = Windows.FirstOrDefault(w => w.Contains(local.DateTime));
        if (window is not null)
            return window.BudgetPercent;
        return OnlyInWindows && Windows.Count > 0 ? null : BudgetPercent;
    }

    internal IEnumerable<ValidationIssue> Validate()
    {
        if (Downloads is < 1 or > 64 || Uploads is < 1 or > 64 || Processors is < 1 or > 64)
            yield return new(IssueSeverity.Error, L.T("performance: паралельність має бути від 1 до 64."));
        if (RequestTimeoutSeconds < 5 || FileTimeoutMinutes < 1 || SimpleUploadMaxMb is < 1 or > 250)
            yield return new(IssueSeverity.Error, L.T("performance: некоректні таймаути або simpleUploadMaxMb (1…250)."));
        if (ExcelProcesses is < 0 or > 16 || ExcelTimeoutMinutes is < 1 or > 120)
            yield return new(IssueSeverity.Error, L.T("performance: процесів Excel від 0 до 16, таймаут Excel від 1 до 120 хв."));
        else if (ExcelProcesses > 0 && ExcelTimeoutMinutes * 3 > FileTimeoutMinutes)
            // A file takes up to three Excel operations (save as xlsm, save back, check); a hang must hit Excel's timeout first.
            yield return new(IssueSeverity.Warning, L.F("performance: таймаут файлу ({0} хв) менший за три таймаути Excel ({1} хв): зависання Excel спрацює як таймаут файлу.", FileTimeoutMinutes, ExcelTimeoutMinutes * 3));
        if (BudgetPercent is < 1 or > 100)
            yield return new(IssueSeverity.Error, L.T("performance.budgetPercent: від 1 до 100."));
        if (OnlyInWindows && Windows.Count == 0)
            yield return new(IssueSeverity.Error, L.T("performance.onlyInWindows потребує хоча б одного вікна."));
        string? zoneError = null;
        try
        {
            ResolveTimeZone();
        }
        catch (Exception ex) when (ex is TimeZoneNotFoundException or InvalidTimeZoneException)
        {
            zoneError = L.F("performance.timeZone: невідомий часовий пояс «{0}».", TimeZone);
        }
        if (zoneError is not null)
            yield return new(IssueSeverity.Error, zoneError);
        for (int i = 0; i < Windows.Count; i++)
        {
            foreach (string error in Windows[i].Validate())
                yield return new(IssueSeverity.Error, $"performance.windows[{i + 1}]: {error}");
        }
    }
}

/// <summary>A weekly time window, e.g. Mon–Fri 20:00–07:00. A window may cross midnight.</summary>
public sealed class ScheduleWindow
{
    /// <summary>Mon, Tue, … Empty — every day. For a window crossing midnight the day is the day it starts.</summary>
    public List<string> Days { get; set; } = [];

    public string From { get; set; } = "00:00";

    public string To { get; set; } = "24:00";

    public int BudgetPercent { get; set; } = 80;

    public bool Contains(DateTime local)
    {
        var from = ParseTime(From);
        var to = ParseTime(To);
        var t = local.TimeOfDay;
        if (from <= to)
            return DayMatches(local.DayOfWeek) && t >= from && t < to;
        // Crosses midnight: the evening part belongs to this day, the morning part to the previous one.
        return (t >= from && DayMatches(local.DayOfWeek)) || (t < to && DayMatches(local.AddDays(-1).DayOfWeek));
    }

    private bool DayMatches(DayOfWeek day) =>
        Days.Count == 0 || Days.Any(d => string.Equals(d.Trim()[..Math.Min(3, d.Trim().Length)], day.ToString()[..3], StringComparison.OrdinalIgnoreCase));

    private static TimeSpan ParseTime(string s) =>
        s == "24:00" ? TimeSpan.FromDays(1) : TimeSpan.ParseExact(s, @"hh\:mm", CultureInfo.InvariantCulture);

    internal IEnumerable<string> Validate()
    {
        foreach (string s in new[] { From, To })
        {
            if (s != "24:00" && !TimeSpan.TryParseExact(s, @"hh\:mm", CultureInfo.InvariantCulture, out _))
                yield return L.F("час «{0}» має бути у форматі ГГ:ХХ.", s);
        }
        string[] known = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
        foreach (string d in Days.Where(d => d.Trim().Length < 3 || !known.Contains(d.Trim()[..3].ToLowerInvariant())))
            yield return L.F("невідомий день «{0}» (Mon, Tue, …).", d);
        if (BudgetPercent is < 1 or > 100)
            yield return L.T("budgetPercent: від 1 до 100.");
    }
}

/// <summary>Levels 1–3 of FEATURES 9.2 and the circuit breaker (RSL-5).</summary>
public sealed class RetryOptions
{
    /// <summary>Level 1: attempts of one HTTP request.</summary>
    public int RequestAttempts { get; set; } = 6;

    /// <summary>Level 2: attempts of one file in the run (transient, conflict).</summary>
    public int FileAttempts { get; set; } = 3;

    /// <summary>Pauses before the 2nd, 3rd… attempt of a file. Seconds allow short values for tests.</summary>
    public List<double> FilePausesMinutes { get; set; } = [1, 5, 15];

    /// <summary>Level 3: how often locked files are tried again.</summary>
    public double LockedWaveMinutes { get; set; } = 10;

    /// <summary>Level 3: how long a file may stay locked before failed:locked.</summary>
    public double LockedLimitMinutes { get; set; } = 120;

    /// <summary>RSL-5: the breaker looks at this many last finished files.</summary>
    public int BreakerWindow { get; set; } = 50;

    /// <summary>RSL-5: share of failures (transient, auth, internal) in the window that stops the run.</summary>
    public double BreakerFailureRatio { get; set; } = 0.5;

    /// <summary>Auth failures in a row that stop the run (mass 401/403).</summary>
    public int AuthFailuresToStop { get; set; } = 10;

    /// <summary>THR-7: only 503 responses for this long — the app is considered blocked and the run stops.</summary>
    public double BlockedMinutes { get; set; } = 15;

    internal IEnumerable<ValidationIssue> Validate()
    {
        if (RequestAttempts is < 1 or > 20 || FileAttempts is < 1 or > 20)
            yield return new(IssueSeverity.Error, L.T("retry: кількість спроб від 1 до 20."));
        if (FilePausesMinutes.Count == 0 || FilePausesMinutes.Any(p => p < 0))
            yield return new(IssueSeverity.Error, L.T("retry.filePausesMinutes: потрібна хоча б одна невід'ємна пауза."));
        if (LockedWaveMinutes <= 0 || LockedLimitMinutes < 0 || BlockedMinutes <= 0)
            yield return new(IssueSeverity.Error, L.T("retry: некоректні інтервали для заблокованих файлів або блокування."));
        if (BreakerWindow < 1 || BreakerFailureRatio is <= 0 or > 1 || AuthFailuresToStop < 1)
            yield return new(IssueSeverity.Error, L.T("retry: некоректні параметри запобіжника."));
    }
}

/// <summary>THR-2: SharePoint Online limits by number of licenses in the tenant (Microsoft docs, 10.2026).</summary>
public enum LicenseTier
{
    UpTo1000,
    UpTo5000,
    UpTo15000,
    UpTo50000,
    Over50000,
}

public static class SpoLimits
{
    /// <summary>RU per minute and per day for one app; RU per 5 minutes for the whole tenant.</summary>
    public static (int AppPerMinute, int AppPerDay, int TenantPer5Min) For(LicenseTier tier) => tier switch
    {
        LicenseTier.UpTo1000 => (1_250, 1_200_000, 18_750),
        LicenseTier.UpTo5000 => (2_500, 2_400_000, 37_500),
        LicenseTier.UpTo15000 => (3_750, 3_600_000, 56_250),
        LicenseTier.UpTo50000 => (5_000, 4_800_000, 75_000),
        _ => (6_250, 6_000_000, 93_750),
    };
}

/// <summary>An app registration in Entra (GUI-26). No secrets: the certificate stays in the Windows store (CFG-6).</summary>
public sealed class Connection
{
    public string Name { get; set; } = string.Empty;

    public string? TenantId { get; set; }

    public string? ClientId { get; set; }

    /// <summary>certificate (app-only, SEC-2) or deviceCode (delegated, fallback).</summary>
    public string AuthMode { get; set; } = "certificate";

    public string? CertificateThumbprint { get; set; }

    /// <summary>CurrentUser or LocalMachine.</summary>
    public string CertificateStore { get; set; } = "CurrentUser";

    public LicenseTier Licenses { get; set; } = LicenseTier.UpTo1000;

    /// <summary>Manual correction of the app limit (THR-2).</summary>
    public int? AppPerMinuteOverride { get; set; }

    public int? AppPerDayOverride { get; set; }

    /// <summary>Manual correction of the tenant limit (RU per 5 minutes) when the tier does not match reality.</summary>
    public int? TenantPer5MinOverride { get; set; }

    public static Connection Parse(string json)
    {
        try
        {
            return JsonSerializer.Deserialize<Connection>(json, Profile.JsonOptions) ?? throw new ProfileException(L.T("Підключення порожнє."));
        }
        catch (JsonException ex)
        {
            throw new ProfileException(L.T("Підключення не розібрано: ") + ex.Message, ex);
        }
    }

    public string ToJson() => JsonSerializer.Serialize(this, Profile.JsonOptions);

    public (int PerMinute, int PerDay, int TenantPer5Min) Limits()
    {
        var l = SpoLimits.For(Licenses);
        return (AppPerMinuteOverride ?? l.AppPerMinute, AppPerDayOverride ?? l.AppPerDay, TenantPer5MinOverride ?? l.TenantPer5Min);
    }

    public IEnumerable<ValidationIssue> Validate()
    {
        string who = L.F("Підключення «{0}»", Name);
        if (string.IsNullOrWhiteSpace(Name))
            yield return new(IssueSeverity.Error, L.T("Підключення без назви."));
        if (!Guid.TryParse(ClientId, out _))
            yield return new(IssueSeverity.Error, L.F("{0}: clientId має бути GUID.", who));
        if (string.IsNullOrWhiteSpace(TenantId))
            yield return new(IssueSeverity.Error, L.F("{0}: не задано tenantId.", who));
        if (AuthMode == "certificate")
        {
            if (string.IsNullOrWhiteSpace(CertificateThumbprint) || !Regex.IsMatch(CertificateThumbprint.Replace(" ", ""), "^[0-9A-Fa-f]{40}$"))
                yield return new(IssueSeverity.Error, L.F("{0}: certificateThumbprint — 40 шістнадцяткових символів.", who));
            if (CertificateStore is not ("CurrentUser" or "LocalMachine"))
                yield return new(IssueSeverity.Error, L.F("{0}: certificateStore — CurrentUser або LocalMachine.", who));
        }
        else if (AuthMode != "deviceCode")
        {
            yield return new(IssueSeverity.Error, L.F("{0}: authMode — certificate або deviceCode.", who));
        }
        if (AppPerMinuteOverride is < 1 || AppPerDayOverride is < 1 || TenantPer5MinOverride is < 1)
            yield return new(IssueSeverity.Error, L.F("{0}: ліміти мають бути додатними.", who));
    }
}

/// <summary>Named set of rules (GUI-22).</summary>
public sealed class Template
{
    public string Name { get; set; } = string.Empty;

    public string? Description { get; set; }

    public List<RuleEntry> Rules { get; set; } = [];

    /// <summary>
    /// Host for server-relative rules when the template is checked on its own (template editor); a run always uses the
    /// profile's baseHost.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? TestBaseHost { get; set; }

    /// <summary>TestBaseHost, else the host of the first absolute rule, else null.</summary>
    public string? GuessBaseHost() =>
        !string.IsNullOrWhiteSpace(TestBaseHost) ? TestBaseHost
        : Rules.SelectMany(r => new[] { r.Old, r.New })
            .Select(u => Uri.TryCreate(u, UriKind.Absolute, out var uri) && uri.Scheme is "http" or "https" ? uri.GetLeftPart(UriPartial.Authority) : null)
            .FirstOrDefault(h => h is not null);

    public static Template Parse(string json)
    {
        try
        {
            return JsonSerializer.Deserialize<Template>(json, Profile.JsonOptions) ?? throw new ProfileException(L.T("Шаблон порожній."));
        }
        catch (JsonException ex)
        {
            throw new ProfileException(L.T("Шаблон не розібрано: ") + ex.Message, ex);
        }
    }

    public string ToJson() => JsonSerializer.Serialize(this, Profile.JsonOptions);
}

public enum FilterVerdict
{
    Accepted,
    Extension,
    Path,
    Date,
    TooLarge,
    TooSmall,
}

/// <summary>Path masks: * — any characters within one folder, ** — any depth, ? — one character. Case-insensitive.</summary>
public static class PathMask
{
    private static readonly Dictionary<string, Regex> Cache = new(StringComparer.Ordinal);

    public static bool IsMatch(string mask, string path)
    {
        Regex regex;
        lock (Cache)
        {
            if (!Cache.TryGetValue(mask, out regex!))
                Cache[mask] = regex = new Regex(ToRegex(mask), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }
        return regex.IsMatch(path.Replace('\\', '/').TrimStart('/'));
    }

    private static string ToRegex(string mask)
    {
        mask = mask.Replace('\\', '/').TrimStart('/');
        var sb = new System.Text.StringBuilder("^");
        for (int i = 0; i < mask.Length; i++)
        {
            char c = mask[i];
            if (c == '*' && i + 1 < mask.Length && mask[i + 1] == '*')
            {
                bool slash = i + 2 < mask.Length && mask[i + 2] == '/';
                sb.Append(slash ? "(?:.*/)?" : ".*");
                i += slash ? 2 : 1;
            }
            else if (c == '*')
            {
                sb.Append("[^/]*");
            }
            else if (c == '?')
            {
                sb.Append("[^/]");
            }
            else
            {
                sb.Append(Regex.Escape(c.ToString()));
            }
        }
        return sb.Append('$').ToString();
    }
}
