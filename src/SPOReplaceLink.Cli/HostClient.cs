using System.Net;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace SPOReplaceLink.Cli;

/// <summary>Thin client of the host REST API: the CLI never runs SharePoint work itself (plan §1).</summary>
internal sealed class HostClient(HttpClient http)
{
    /// <param name="key">Key of the instance (api.json in its data folder); without it the host answers 401.</param>
    public static HostClient Create(string baseUrl, HttpMessageHandler? handler = null, string? key = null)
    {
        var client = handler is null ? new HttpClient() : new HttpClient(handler);
        client.BaseAddress = new Uri(baseUrl.TrimEnd('/') + "/");
        client.Timeout = TimeSpan.FromMinutes(2);
        if (key is not null)
            client.DefaultRequestHeaders.Add(Core.Hosting.ApiEndpoint.KeyHeader, key);
        // The host answers in the language of this CLI.
        client.DefaultRequestHeaders.AcceptLanguage.Add(new System.Net.Http.Headers.StringWithQualityHeaderValue(Core.Localization.L.Language));
        return new HostClient(client);
    }

    public async Task<JsonNode?> GetAsync(string path, CancellationToken ct = default) =>
        await Send(new HttpRequestMessage(HttpMethod.Get, path), ct);

    public async Task<string> GetTextAsync(string path, CancellationToken ct = default)
    {
        using var response = await SendRaw(new HttpRequestMessage(HttpMethod.Get, path), ct);
        return await response.Content.ReadAsStringAsync(ct);
    }

    public async Task<Stream> GetStreamAsync(string path, CancellationToken ct = default)
    {
        var response = await SendRaw(new HttpRequestMessage(HttpMethod.Get, path), ct);
        return await response.Content.ReadAsStreamAsync(ct);
    }

    public Task<JsonNode?> PostAsync(string path, object? body = null, CancellationToken ct = default) =>
        Send(new HttpRequestMessage(HttpMethod.Post, path) { Content = body is null ? null : JsonContent.Create(body, options: JsonOptions) }, ct);

    public Task<JsonNode?> PutJsonTextAsync(string path, string json, CancellationToken ct = default) =>
        Send(new HttpRequestMessage(HttpMethod.Put, path) { Content = new StringContent(json, Encoding.UTF8, "application/json") }, ct);

    public Task<JsonNode?> DeleteAsync(string path, CancellationToken ct = default) =>
        Send(new HttpRequestMessage(HttpMethod.Delete, path), ct);

    public static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    private async Task<JsonNode?> Send(HttpRequestMessage request, CancellationToken ct)
    {
        using var response = await SendRaw(request, ct);
        string text = await response.Content.ReadAsStringAsync(ct);
        return text.Length == 0 ? null : JsonNode.Parse(text);
    }

    private async Task<HttpResponseMessage> SendRaw(HttpRequestMessage request, CancellationToken ct)
    {
        HttpResponseMessage response;
        try
        {
            response = await http.SendAsync(request, ct);
        }
        catch (HttpRequestException ex)
        {
            throw new HostException(Core.Localization.L.F("Хост недоступний за адресою {0}: {1}. Запустіть SPOReplaceLink.exe під цим обліковим записом (CLI працює, поки відкрите його вікно), або вкажіть --data/--host.", http.BaseAddress, ex.Message));
        }
        if (response.IsSuccessStatusCode)
            return response;

        string text = await response.Content.ReadAsStringAsync(ct);
        response.Dispose();
        string message = text;
        var issues = new List<string>();
        try
        {
            var node = JsonNode.Parse(text);
            message = node?["error"]?.GetValue<string>() ?? text;
            foreach (var issue in node?["issues"]?.AsArray() ?? [])
                issues.Add($"{issue!["severity"]}: {issue["message"]}");
        }
        catch (JsonException)
        {
        }
        throw new HostException(message, response.StatusCode, issues);
    }
}

internal sealed class HostException(string message, HttpStatusCode? status = null, IReadOnlyList<string>? issues = null) : Exception(message)
{
    public HttpStatusCode? Status { get; } = status;

    public IReadOnlyList<string> Issues { get; } = issues ?? [];
}
