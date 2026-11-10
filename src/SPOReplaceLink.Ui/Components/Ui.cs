using System.Globalization;
using MudBlazor;
using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.State;

namespace SPOReplaceLink.Ui.Components;

/// <summary>Labels and colors shared by the pages.</summary>
public static class Ui
{
    /// <summary>Numbers and dates in the culture of the current language.</summary>
    public static CultureInfo Culture => L.Culture;

    public static string N(long n) => n.ToString("N0", Culture);

    /// <summary>Date and time format of the current language.</summary>
    public static string TimeFormat => L.Language == L.Ukrainian ? "dd.MM.yyyy HH:mm:ss" : "yyyy-MM-dd HH:mm:ss";

    public static string Time(DateTimeOffset? t) => t is { } v ? v.ToLocalTime().ToString(TimeFormat, Culture) : "—";

    public static string Duration(TimeSpan? t) => t is not { } d ? "—"
        : d.TotalDays >= 1 ? L.F("{0} д {1:hh\\:mm}", (int)d.TotalDays, d)
        : d.TotalHours >= 1 ? d.ToString(@"h\:mm\:ss", Culture)
        : d.ToString(@"m\:ss", Culture);

    public static string Size(long bytes) => bytes switch
    {
        >= 1024L * 1024 * 1024 => L.F("{0:0.0} ГБ", bytes / 1024.0 / 1024 / 1024),
        >= 1024 * 1024 => L.F("{0:0.0} МБ", bytes / 1024.0 / 1024),
        >= 1024 => L.F("{0:0} КБ", bytes / 1024.0),
        _ => L.F("{0} Б", bytes),
    };

    public static string Mode(string mode) => mode switch
    {
        RunModes.Inventory => L.T("inventory — перелік"),
        RunModes.Scan => L.T("scan — пробний"),
        RunModes.Apply => L.T("apply — запис"),
        RunModes.Verify => L.T("verify — перевірка"),
        _ => mode,
    };

    public static string RunStatus(string s) => s switch
    {
        RunStatuses.Created => L.T("створено"),
        RunStatuses.Running => L.T("виконується"),
        RunStatuses.Paused => L.T("пауза"),
        RunStatuses.Stopped => L.T("зупинено"),
        RunStatuses.Aborted => L.T("перервано"),
        RunStatuses.Completed => L.T("завершено"),
        RunStatuses.CompletedWithErrors => L.T("завершено з помилками"),
        _ => s,
    };

    public static Color RunColor(string s) => s switch
    {
        RunStatuses.Running => Color.Info,
        RunStatuses.Paused => Color.Warning,
        RunStatuses.Completed => Color.Success,
        RunStatuses.CompletedWithErrors or RunStatuses.Aborted => Color.Error,
        _ => Color.Default,
    };

    public static Color FileColor(string? code) => code switch
    {
        null => Color.Default,
        _ when code.StartsWith("failed:", StringComparison.Ordinal) => Color.Error,
        _ when code.EndsWith(":residual", StringComparison.Ordinal) || code.Contains("needs_excel", StringComparison.Ordinal) || code == FileStatuses.ChangesPending => Color.Warning,
        FileStatuses.Done or FileStatuses.Verified => Color.Success,
        FileStatuses.Planned => Color.Info,
        FileStatuses.Deferred or FileStatuses.Pending or FileStatuses.InProgress => Color.Secondary,
        _ => Color.Default,
    };

    public static string ErrorClass(string? c) => c switch
    {
        "transient" => L.T("тимчасова"),
        "conflict" => L.T("конфлікт"),
        "locked" => L.T("заблоковано"),
        "excel" => "Excel",
        "auth" => L.T("доступ"),
        "file" => L.T("файл"),
        "validation" => L.T("самоперевірка"),
        "internal" => L.T("внутрішня"),
        null or "" => "—",
        _ => c,
    };

    public static string LinkKind(string kind) => kind switch
    {
        "Hyperlink" => L.T("гіперпосилання"),
        "LinkSource" => "Link Sources",
        "Formula" => L.T("формули"),
        "CachedValue" => L.T("кеш формул"),
        "OtherExternal" => L.T("інші зовнішні"),
        "WorkbookPath" => L.T("шлях збереження"),
        _ => kind,
    };

    /// <summary>Status filter values; "attention" is not a status but all files the "Потребують уваги" card counts.</summary>
    public static readonly string[] FileStatusFilters =
    [
        FileStatuses.Attention, "pending", "in_progress", "deferred", "done", "unchanged", "planned", "verified", "changes_pending", "needs_excel",
        "ignored", "skipped", "skipped:encrypted", "failed", "failed:transient", "failed:conflict", "failed:locked", "failed:excel", "failed:auth", "failed:file",
        "failed:validation", "failed:internal", "done:residual", "unchanged:residual", "planned:residual", "verified:residual",
    ];

    public static string FileStatusFilter(string value) =>
        value == FileStatuses.Attention ? L.T("attention — потребують уваги") : value;

    public static readonly string[] ErrorClasses = ["transient", "conflict", "locked", "excel", "auth", "file", "validation", "internal"];
}
