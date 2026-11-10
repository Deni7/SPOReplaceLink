using System.Net;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Spo;

/// <summary>A Graph request that failed after all request-level retries (RTY-1: class, HTTP code, request-id).</summary>
public class SpoException(
    ErrorClass errorClass,
    string message,
    HttpStatusCode? status = null,
    string? code = null,
    string? requestId = null,
    Exception? inner = null) : Exception(message, inner)
{
    public ErrorClass ErrorClass { get; } = errorClass;

    public HttpStatusCode? Status { get; } = status;

    /// <summary>Graph error code, e.g. resourceLocked, itemNotFound.</summary>
    public string? Code { get; } = code;

    /// <summary>request-id from the response, for Microsoft support.</summary>
    public string? RequestId { get; } = requestId;

    public bool IsNotFound => Status == HttpStatusCode.NotFound;

    public static ErrorClass Classify(HttpStatusCode status, string? code) => (int)status switch
    {
        401 or 403 => ErrorClass.Auth,
        409 or 412 => code is not null && code.Contains("lock", StringComparison.OrdinalIgnoreCase) ? ErrorClass.Locked : ErrorClass.Conflict,
        423 => ErrorClass.Locked,
        408 or 429 or >= 500 => ErrorClass.Transient,
        413 => ErrorClass.File,
        _ when code is not null && (code.Contains("lock", StringComparison.OrdinalIgnoreCase) || code.Contains("checkedOut", StringComparison.OrdinalIgnoreCase)) => ErrorClass.Locked,
        _ => ErrorClass.Internal,
    };
}

/// <summary>THR-7: the app gets only 503 for a long time — SharePoint has probably blocked it. The run must stop.</summary>
public sealed class SpoBlockedException(string app, TimeSpan duration)
    : SpoException(ErrorClass.Transient,
        L.F("Можливе блокування застосунку «{0}» у SharePoint: лише 503 протягом {1:0} хв. Прогін зупинено; подальші спроби лише подовжують блокування.", app, duration.TotalMinutes))
{
    public string App { get; } = app;
}
