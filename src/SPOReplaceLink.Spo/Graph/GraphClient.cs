using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Spo.Auth;
using SPOReplaceLink.Spo.Throttling;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Spo.Graph;

public sealed class GraphClientOptions
{
    public string BaseUrl { get; init; } = "https://graph.microsoft.com/v1.0";

    /// <summary>RSL-6: identifies the traffic as a known internal app.</summary>
    public string UserAgent { get; init; } = "NONISV|SPOReplaceLink|SPOReplaceLink/" + (typeof(GraphClient).Assembly.GetName().Version?.ToString(2) ?? "1.0");

    /// <summary>Level 1 of FEATURES 9.2.</summary>
    public int RequestAttempts { get; init; } = 6;

    /// <summary>Time to get response headers; transfers of file content are bounded by the caller's token.</summary>
    public TimeSpan RequestTimeout { get; init; } = TimeSpan.FromSeconds(120);

    public int ChunkSize { get; init; } = 10 * 1024 * 1024; // multiple of 320 KiB

    /// <summary>Upper bound of the exponential pause between retries.</summary>
    public TimeSpan MaxBackoff { get; init; } = TimeSpan.FromSeconds(60);

    /// <summary>First pause of the exponential backoff (tests use milliseconds).</summary>
    public TimeSpan BaseBackoff { get; init; } = TimeSpan.FromSeconds(1);
}

/// <summary>
/// Graph client of one app. Every request goes through the app budget, the tenant budget and the app throttle
/// state; failures are retried here (level 1) and surface as <see cref="SpoException"/> with a class (RTY-1).
/// </summary>
public sealed class GraphClient
{
    private readonly HttpClient _http;
    private readonly ITokenProvider _tokens;
    private readonly RuBudget _appBudget;
    private readonly RuBudget? _tenantBudget;
    private readonly ThrottleState _throttle;
    private readonly GraphClientOptions _options;
    private readonly ILogger _logger;
    private readonly TimeProvider _time;

    public GraphClient(string app, HttpClient http, ITokenProvider tokens, RuBudget appBudget, RuBudget? tenantBudget,
        ThrottleState throttle, GraphClientOptions options, ILogger logger, TimeProvider? time = null)
    {
        App = app;
        _http = http;
        _tokens = tokens;
        _appBudget = appBudget;
        _tenantBudget = tenantBudget;
        _throttle = throttle;
        _options = options;
        _logger = logger;
        _time = time ?? TimeProvider.System;
    }

    public string App { get; }

    // ---- operations ----

    /// <summary>Only gets a token: proves the app registration and the certificate work.</summary>
    public Task CheckTokenAsync(CancellationToken ct) => _tokens.GetTokenAsync(forceRefresh: false, ct);

    /// <summary>Site by URL: GET /sites/{host}:/{path}. Works with Sites.Selected, unlike search.</summary>
    public Task<SpoSite> GetSiteAsync(string siteUrl, CancellationToken ct)
    {
        var uri = new Uri(siteUrl.TrimEnd('/'));
        string path = uri.AbsolutePath.Trim('/');
        string url = path.Length == 0 ? $"{_options.BaseUrl}/sites/{uri.Host}" : $"{_options.BaseUrl}/sites/{uri.Host}:/{EscapePath(Uri.UnescapeDataString(path))}";
        return SendJsonAsync(() => new HttpRequestMessage(HttpMethod.Get, url + "?$select=id,webUrl"), RequestCost.Get,
            e => new SpoSite(e.GetProperty("id").GetString()!, e.GetProperty("webUrl").GetString()!), "site", ct);
    }

    public Task<IReadOnlyList<SpoDrive>> GetDrivesAsync(string siteId, CancellationToken ct) =>
        SendJsonAsync<IReadOnlyList<SpoDrive>>(() => new HttpRequestMessage(HttpMethod.Get, $"{_options.BaseUrl}/sites/{siteId}/drives?$select=id,name,webUrl"),
            RequestCost.List,
            e => e.GetProperty("value").EnumerateArray()
                .Select(d => new SpoDrive(d.GetProperty("id").GetString()!, d.GetProperty("name").GetString()!, d.GetProperty("webUrl").GetString()!))
                .ToList(),
            "drives", ct);

    /// <summary>GUI-17: site search. With Sites.Selected it may return nothing — entering the URL is the main way.</summary>
    public Task<IReadOnlyList<SpoSite>> SearchSitesAsync(string query, CancellationToken ct) =>
        SendJsonAsync<IReadOnlyList<SpoSite>>(() => new HttpRequestMessage(HttpMethod.Get, $"{_options.BaseUrl}/sites?search={Uri.EscapeDataString(query)}&$select=id,webUrl,displayName"),
            RequestCost.List,
            e => e.GetProperty("value").EnumerateArray().Select(s => new SpoSite(s.GetProperty("id").GetString()!, s.GetProperty("webUrl").GetString()!)).ToList(),
            "site-search", ct);

    /// <summary>GUI-17: subfolders of a folder (null — the library root) for the folder tree.</summary>
    public async Task<IReadOnlyList<SpoItem>> GetChildFoldersAsync(string driveId, string? itemId, CancellationToken ct)
    {
        var result = new List<SpoItem>();
        string? url = $"{_options.BaseUrl}/drives/{driveId}/{(itemId is null ? "root" : "items/" + itemId)}/children?$select=id,name,folder,parentReference&$top=999";
        while (url is not null)
        {
            string current = url;
            var (items, next) = await SendJsonAsync(() => new HttpRequestMessage(HttpMethod.Get, current), RequestCost.List,
                e => (e.GetProperty("value").EnumerateArray().Select(SpoItem.Parse).ToList(),
                      e.TryGetProperty("@odata.nextLink", out var n) ? n.GetString() : null),
                "children", ct).ConfigureAwait(false);
            result.AddRange(items.Where(i => i.IsFolder));
            url = next;
        }
        return result.OrderBy(i => i.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
    }

    public Task<SpoItem> GetItemByPathAsync(string driveId, string path, CancellationToken ct) =>
        SendJsonAsync(() => new HttpRequestMessage(HttpMethod.Get, $"{_options.BaseUrl}/drives/{driveId}/root:/{EscapePath(path.Trim('/'))}"),
            RequestCost.Get, SpoItem.Parse, "folder", ct);

    /// <summary>ENU-1: a page of root delta; <paramref name="link"/> is the saved nextLink to continue from.</summary>
    public Task<DeltaPage> GetDeltaPageAsync(string driveId, string? link, CancellationToken ct)
    {
        string url = link ?? $"{_options.BaseUrl}/drives/{driveId}/root/delta?$select=id,name,size,file,folder,root,deleted,parentReference,eTag,cTag,lastModifiedDateTime,lastModifiedBy";
        return SendJsonAsync(() => new HttpRequestMessage(HttpMethod.Get, url), link is null ? RequestCost.List : RequestCost.DeltaWithToken,
            e => new DeltaPage(
                e.GetProperty("value").EnumerateArray().Select(SpoItem.Parse).ToList(),
                e.TryGetProperty("@odata.nextLink", out var n) ? n.GetString() : null,
                e.TryGetProperty("@odata.deltaLink", out var d) ? d.GetString() : null),
            "delta", ct);
    }

    /// <summary>Fresh metadata with the current version (UPL-4) and check-out state (UPL-6).</summary>
    public Task<SpoItem> GetItemAsync(string driveId, string itemId, CancellationToken ct) =>
        SendJsonAsync(() => new HttpRequestMessage(HttpMethod.Get,
                $"{_options.BaseUrl}/drives/{driveId}/items/{itemId}?$select=id,name,size,eTag,cTag,lastModifiedDateTime,lastModifiedBy,parentReference,file,folder,publication&$expand=listItem($select=id;$expand=fields($select=_UIVersionString,CheckoutUser))"),
            RequestCost.Get, SpoItem.Parse, "metadata", ct);

    /// <summary>Downloads file content to <paramref name="destination"/> (follows the redirect to the download URL).</summary>
    public Task DownloadAsync(string driveId, string itemId, string destination, CancellationToken ct) =>
        SendAsync(() => new HttpRequestMessage(HttpMethod.Get, $"{_options.BaseUrl}/drives/{driveId}/items/{itemId}/content"),
            RequestCost.Download, authenticated: true, async response =>
            {
                await using var file = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);
                await response.Content.CopyToAsync(file, ct).ConfigureAwait(false);
                return true;
            }, "download", ct);

    /// <summary>GUI-6: content of an earlier version (e.g. "3.0" — the one before the tool wrote the file).</summary>
    public Task DownloadVersionAsync(string driveId, string itemId, string versionId, string destination, CancellationToken ct) =>
        SendAsync(() => new HttpRequestMessage(HttpMethod.Get, $"{_options.BaseUrl}/drives/{driveId}/items/{itemId}/versions/{Uri.EscapeDataString(versionId)}/content"),
            RequestCost.Download, authenticated: true, async response =>
            {
                await using var file = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);
                await response.Content.CopyToAsync(file, ct).ConfigureAwait(false);
                return true;
            }, "download-version", ct);

    /// <summary>
    /// UPL-2, UPL-3: replaces the content only if the item still has <paramref name="eTag"/>;
    /// otherwise 412 → <see cref="ErrorClass.Conflict"/>. Large files go through an upload session.
    /// </summary>
    public async Task<SpoItem> UploadAsync(string driveId, string itemId, string source, string eTag, long simpleMaxBytes, CancellationToken ct)
    {
        long size = new FileInfo(source).Length;
        if (size <= simpleMaxBytes)
        {
            return await SendJsonAsync(() =>
            {
                var request = new HttpRequestMessage(HttpMethod.Put, $"{_options.BaseUrl}/drives/{driveId}/items/{itemId}/content")
                {
                    Content = new StreamContent(File.OpenRead(source)),
                };
                request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                request.Headers.TryAddWithoutValidation("If-Match", eTag);
                return request;
            }, RequestCost.Upload, SpoItem.Parse, "upload", ct, TransferTimeout(size)).ConfigureAwait(false);
        }
        return await UploadInChunksAsync(driveId, itemId, source, size, eTag, ct).ConfigureAwait(false);
    }

    private async Task<SpoItem> UploadInChunksAsync(string driveId, string itemId, string source, long size, string eTag, CancellationToken ct)
    {
        string uploadUrl = await SendJsonAsync(() =>
        {
            var request = new HttpRequestMessage(HttpMethod.Post, $"{_options.BaseUrl}/drives/{driveId}/items/{itemId}/createUploadSession")
            {
                Content = new StringContent("""{"item":{"@microsoft.graph.conflictBehavior":"replace"}}""", Encoding.UTF8, "application/json"),
            };
            request.Headers.TryAddWithoutValidation("If-Match", eTag);
            return request;
        }, RequestCost.Upload, e => e.GetProperty("uploadUrl").GetString()!, "upload-session", ct).ConfigureAwait(false);

        long offset = 0;
        int resumes = 0;
        while (true)
        {
            long start = offset;
            int length = (int)Math.Min(_options.ChunkSize, size - start);
            byte[] chunk = new byte[length];
            await using (var fs = File.OpenRead(source))
            {
                fs.Seek(start, SeekOrigin.Begin);
                await fs.ReadExactlyAsync(chunk, ct).ConfigureAwait(false);
            }

            // If-Match is only checked when the session is created: before the chunk that commits the file,
            // make sure nobody saved it meanwhile, otherwise their edit would be overwritten.
            if (start + length >= size)
            {
                var current = await GetItemAsync(driveId, itemId, ct).ConfigureAwait(false);
                if (current.ETag != eTag)
                {
                    try
                    {
                        await SendAsync(() => new HttpRequestMessage(HttpMethod.Delete, uploadUrl), RequestCost.Get, authenticated: false,
                            _ => Task.FromResult(true), "upload-cancel", ct).ConfigureAwait(false);
                    }
                    catch (SpoException)
                    {
                    }
                    throw new SpoException(ErrorClass.Conflict, L.T("Файл змінили під час вивантаження частинами; сесію скасовано."), HttpStatusCode.PreconditionFailed);
                }
            }

            (SpoItem? Item, long Next) step;
            try
            {
                step = await SendAsync(() =>
                {
                    var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl) { Content = new ByteArrayContent(chunk) };
                    request.Content.Headers.ContentRange = new ContentRangeHeaderValue(start, start + length - 1, size);
                    return request;
                }, RequestCost.UploadChunk, authenticated: false, async response =>
                {
                    using var doc = await ReadJsonAsync(response, ct).ConfigureAwait(false);
                    if (response.StatusCode is HttpStatusCode.OK or HttpStatusCode.Created)
                        return ((SpoItem?)SpoItem.Parse(doc.RootElement), size);
                    return ((SpoItem?)null, NextExpected(doc.RootElement, start + length));
                }, "upload-chunk", ct, TransferTimeout(length)).ConfigureAwait(false);
            }
            catch (SpoException ex) when (ex.ErrorClass == ErrorClass.Transient && resumes++ < 3)
            {
                // UPL-3: ask the session where to continue instead of starting over.
                _logger.LogWarning(L.T("Відновлення сесії вивантаження {Item} після помилки: {Message}"), itemId, ex.Message);
                offset = await SendAsync(() => new HttpRequestMessage(HttpMethod.Get, uploadUrl), RequestCost.Get, authenticated: false,
                    async response =>
                    {
                        using var doc = await ReadJsonAsync(response, ct).ConfigureAwait(false);
                        return NextExpected(doc.RootElement, start);
                    }, "upload-status", ct).ConfigureAwait(false);
                continue;
            }

            if (step.Item is not null)
                return step.Item;
            offset = step.Next;
        }
    }

    private static long NextExpected(JsonElement e, long fallback)
    {
        if (e.TryGetProperty("nextExpectedRanges", out var ranges) && ranges.GetArrayLength() > 0)
        {
            string first = ranges[0].GetString()!;
            return long.Parse(first.Split('-')[0], System.Globalization.CultureInfo.InvariantCulture);
        }
        return fallback;
    }

    // ---- transport ----

    private Task<T> SendJsonAsync<T>(Func<HttpRequestMessage> factory, int cost, Func<JsonElement, T> parse, string operation, CancellationToken ct,
        TimeSpan? timeout = null) =>
        SendAsync(factory, cost, authenticated: true, async response =>
        {
            using var doc = await ReadJsonAsync(response, ct).ConfigureAwait(false);
            return parse(doc.RootElement);
        }, operation, ct, timeout);

    /// <summary>The request body is sent within the timeout too: allow for a slow link (128 KB/s) on top of the base time.</summary>
    private TimeSpan TransferTimeout(long bytes) => _options.RequestTimeout + TimeSpan.FromSeconds(bytes / (128.0 * 1024));

    private static async Task<JsonDocument> ReadJsonAsync(HttpResponseMessage response, CancellationToken ct)
    {
        await using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        return await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
    }

    /// <summary>
    /// RSL-1…RSL-3, THR-1, THR-3: one request with budget, throttle state, token refresh on 401,
    /// Retry-After on 429/503 for the whole app, exponential backoff with jitter for network errors and 5xx.
    /// </summary>
    private async Task<T> SendAsync<T>(Func<HttpRequestMessage> factory, int cost, bool authenticated,
        Func<HttpResponseMessage, Task<T>> read, string operation, CancellationToken ct, TimeSpan? timeout = null)
    {
        var limit = timeout ?? _options.RequestTimeout;
        bool refreshed = false;
        SpoException? last = null;
        for (int attempt = 1; attempt <= _options.RequestAttempts; attempt++)
        {
            if (_throttle.IsBlocked)
                throw new SpoBlockedException(App, _throttle.Only503For);

            await _appBudget.WaitAsync(cost, ct).ConfigureAwait(false);
            if (_tenantBudget is not null)
                await _tenantBudget.WaitAsync(cost, ct).ConfigureAwait(false);
            await _throttle.EnterAsync(ct).ConfigureAwait(false);

            HttpResponseMessage? response = null;
            TimeSpan? delay = null;
            try
            {
                using var request = factory();
                request.Headers.UserAgent.Clear();
                request.Headers.TryAddWithoutValidation("User-Agent", _options.UserAgent);
                if (authenticated)
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await _tokens.GetTokenAsync(forceRefresh: false, ct).ConfigureAwait(false));

                using (var deadline = CancellationTokenSource.CreateLinkedTokenSource(ct))
                {
                    deadline.CancelAfter(limit);
                    try
                    {
                        response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token).ConfigureAwait(false);
                    }
                    catch (OperationCanceledException) when (!ct.IsCancellationRequested)
                    {
                        throw new TimeoutException(L.F("Немає відповіді за {0:0} с.", limit.TotalSeconds));
                    }
                }

                int code = (int)response.StatusCode;
                if (code is >= 200 and < 300)
                {
                    _throttle.OnSuccess();
                    return await read(response).ConfigureAwait(false);
                }

                var (errorCode, message) = await ReadErrorAsync(response, ct).ConfigureAwait(false);
                string? requestId = RequestId(response);

                if (code == 401 && authenticated && !refreshed)
                {
                    _throttle.OnOtherResponse();
                    refreshed = true;
                    await _tokens.GetTokenAsync(forceRefresh: true, ct).ConfigureAwait(false);
                    _logger.LogInformation(L.T("{App}: 401 на {Operation}, токен оновлено примусово"), App, operation);
                    attempt--; // the refresh itself is not a failed attempt
                    continue;
                }

                if (code is 429 or 503)
                {
                    var retryAfter = RetryAfter(response) ?? Backoff(attempt, TimeSpan.FromSeconds(10));
                    _throttle.OnThrottled(retryAfter, code == 503);
                    _logger.LogWarning(L.T("{App}: {Code} на {Operation}, пауза застосунку {Seconds:0} с"), App, code, operation, retryAfter.TotalSeconds);
                    last = new SpoException(ErrorClass.Transient, $"{code} {errorCode}: {message}", response.StatusCode, errorCode, requestId);
                    continue; // the pause is applied by the throttle state for all requests
                }

                _throttle.OnOtherResponse();
                var errorClass = SpoException.Classify(response.StatusCode, errorCode);
                var ex = new SpoException(errorClass, $"{code} {errorCode}: {message}", response.StatusCode, errorCode, requestId);
                if (errorClass != ErrorClass.Transient)
                    throw ex;
                last = ex;
                delay = Backoff(attempt, _options.BaseBackoff);
            }
            catch (Exception ex) when (ex is HttpRequestException or TimeoutException or IOException or JsonException)
            {
                // JsonException: a body cut off in transit.
                last = new SpoException(ErrorClass.Transient, $"{operation}: {ex.Message}", inner: ex);
                delay = Backoff(attempt, _options.BaseBackoff);
            }
            finally
            {
                response?.Dispose();
                _throttle.Exit();
            }

            if (delay is { } d && attempt < _options.RequestAttempts)
                await Task.Delay(d, _time, ct).ConfigureAwait(false);
        }
        throw last ?? new SpoException(ErrorClass.Transient, L.F("{0}: вичерпано спроби.", operation));
    }

    private TimeSpan Backoff(int attempt, TimeSpan first)
    {
        double ms = first.TotalMilliseconds * Math.Pow(2, attempt - 1);
        ms = Math.Min(ms, _options.MaxBackoff.TotalMilliseconds);
        return TimeSpan.FromMilliseconds(ms * (0.75 + (Random.Shared.NextDouble() * 0.5)));
    }

    private TimeSpan? RetryAfter(HttpResponseMessage response)
    {
        var header = response.Headers.RetryAfter;
        if (header?.Delta is { } delta)
            return delta;
        if (header?.Date is { } date)
        {
            var wait = date - _time.GetUtcNow();
            return wait > TimeSpan.Zero ? wait : TimeSpan.Zero;
        }
        return null;
    }

    private static string? RequestId(HttpResponseMessage response) =>
        response.Headers.TryGetValues("request-id", out var values) ? values.FirstOrDefault()
        : response.Headers.TryGetValues("SPRequestGuid", out var sp) ? sp.FirstOrDefault() : null;

    private static async Task<(string? Code, string Message)> ReadErrorAsync(HttpResponseMessage response, CancellationToken ct)
    {
        string body;
        try
        {
            body = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is HttpRequestException or IOException)
        {
            return (null, response.ReasonPhrase ?? string.Empty);
        }
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("error", out var error))
            {
                string? code = error.TryGetProperty("code", out var c) ? c.GetString() : null;
                if (error.TryGetProperty("innerError", out var inner) && inner.TryGetProperty("code", out var ic) && ic.GetString() is { Length: > 0 } innerCode)
                    code = innerCode;
                return (code, error.TryGetProperty("message", out var m) ? m.GetString() ?? string.Empty : string.Empty);
            }
        }
        catch (JsonException)
        {
        }
        return (null, body.Length > 300 ? body[..300] : body);
    }

    /// <summary>Escapes each segment of a path for a Graph URL.</summary>
    internal static string EscapePath(string path) =>
        string.Join('/', path.Split('/', StringSplitOptions.RemoveEmptyEntries).Select(Uri.EscapeDataString));
}
