using System.Text.Json;
using System.Text.Json.Serialization;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.ExcelWorker;

/// <summary>
/// Exchange between the host and a worker process (XLS-1): one JSON line per message over stdin/stdout.
/// The worker starts with a <see cref="WorkerReady"/> line, then answers each request with one response line.
/// </summary>
public static class Protocol
{
    public static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
    };

    /// <summary>What the file result says when a shared workbook could only be saved with sharing turned off.</summary>
    public static string UnsharedWarning =>
        L.T("Книга мала застарілий спільний доступ, і Excel не зміг зберегти її журнал редакцій: спільний доступ знято, журнал змін не збережено.");
}

/// <summary>Excel file formats (XlFileFormat).</summary>
public static class XlFormat
{
    public const int Xlsb = 50;
    public const int Xlsx = 51;
    public const int Xlsm = 52;
    public const int Xls = 56;

    public static int? ForExtension(string ext) => ext.ToLowerInvariant() switch
    {
        ".xls" => Xls,
        ".xlsb" => Xlsb,
        ".xlsm" => Xlsm,
        ".xlsx" => Xlsx,
        _ => null,
    };
}

/// <summary>Startup lines: first with the Excel pid as soon as Excel runs (Ready false, no error), then Ready or an error.</summary>
public sealed record WorkerReady(bool Ready, int ExcelPid, string? Error = null, ExcelErrorKind? Kind = null);

/// <param name="Op">convert — open Input and save it as Output in Format; inspect — report the links Excel sees in Input; quit.</param>
/// <param name="Inspect">convert: also report the links Excel holds when it saves.</param>
public sealed record WorkerRequest(string Op, string? Input = null, string? Output = null, int Format = 0, bool Inspect = false);

/// <param name="Restart">Excel showed a dialog during the request: the host closes this worker and its Excel.</param>
/// <param name="Warning">convert: the file was saved, but not quite as it was (e.g. sharing turned off).</param>
public sealed record WorkerResponse(
    bool Ok,
    string? Error = null,
    ExcelErrorKind? Kind = null,
    List<string>? LinkSources = null,
    List<string>? Hyperlinks = null,
    List<string>? Formulas = null,
    int? SourceFormat = null,
    bool? Restart = null,
    string? Warning = null);

public enum ExcelErrorKind
{
    /// <summary>Excel is not installed or the worker cannot start: files wait as needs_excel.</summary>
    Unavailable,
    /// <summary>No answer within the timeout: the processes were killed.</summary>
    Timeout,
    /// <summary>The worker or Excel exited in the middle of a request.</summary>
    Crashed,
    /// <summary>The file has a password to open.</summary>
    Password,
    /// <summary>Excel cannot open or save this file.</summary>
    File,
    /// <summary>Any other Excel failure.</summary>
    Excel,
}

public sealed class ExcelException(ExcelErrorKind kind, string message, Exception? inner = null) : Exception(message, inner)
{
    public ExcelErrorKind Kind { get; } = kind;
}
