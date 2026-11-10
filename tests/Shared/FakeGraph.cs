using System.Collections.Concurrent;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Spo;
using SPOReplaceLink.Spo.Auth;
using SPOReplaceLink.Spo.Graph;
using SPOReplaceLink.Spo.Throttling;

namespace SPOReplaceLink.Tests.Shared;

/// <summary>
/// In-memory emulation of the Graph endpoints the tool uses, with eTags, versions, check-out, locks,
/// delta paging, upload sessions and fault injection. Requests are attributed to apps by their token.
/// </summary>
public sealed class FakeGraph : HttpMessageHandler
{
    public const string Base = "https://graph.test/v1.0";
    public const string Host = "contoso.sharepoint.com";
    private const string UploadBase = "https://upload.test/session/";

    private readonly object _sync = new();
    private readonly Dictionary<string, (string Id, string WebUrl)> _sites = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, List<FakeDrive>> _drivesBySite = new();
    private readonly Dictionary<string, FakeDrive> _drives = new();
    private readonly Dictionary<string, Session> _sessions = new();
    private readonly ConcurrentDictionary<string, int> _tokenGeneration = new();
    private int _nextId = 1;

    public int DeltaPageSize { get; set; } = 3;

    /// <summary>Delta returns files before their folders (allowed by Graph; tests path resolution).</summary>
    public bool FilesBeforeFolders { get; set; }

    /// <summary>Called for every request before normal handling; a non-null response is returned as is.</summary>
    public List<Func<FakeRequest, HttpResponseMessage?>> Faults { get; } = [];

    /// <summary>Called after a successful upload of an item's content (before the response is sent).</summary>
    public Action<FakeRequest, FakeItem>? OnUploaded { get; set; }

    /// <summary>Called after content was served.</summary>
    public Action<FakeRequest, FakeItem>? OnDownloaded { get; set; }

    public ConcurrentQueue<FakeRequest> Log { get; } = new();

    public sealed class FakeDrive
    {
        public required string Id { get; init; }
        public required string Name { get; init; }
        public required string WebUrl { get; init; }
        public Dictionary<string, FakeItem> Items { get; } = [];
        public required FakeItem Root { get; init; }
    }

    public sealed class FakeItem
    {
        public required string Id { get; init; }
        public required string Name { get; set; }
        public string? ParentId { get; set; }
        public bool IsFolder { get; init; }
        public bool IsRoot { get; init; }
        public byte[] Content { get; set; } = [];
        public int Version { get; set; } = 1;
        public string ETag => $"\"{{{Id}}},{Version}\"";
        public string? CheckedOutBy { get; set; }
        /// <summary>Uploads answered with 423 before the lock goes away.</summary>
        public int LockedUploads { get; set; }
        public int Uploads { get; set; }
        /// <summary>Content of earlier versions by major version number.</summary>
        [System.Text.Json.Serialization.JsonObjectCreationHandling(System.Text.Json.Serialization.JsonObjectCreationHandling.Populate)]
        public Dictionary<int, byte[]> History { get; } = [];
        public DateTimeOffset Modified { get; set; } = new(2024, 1, 1, 0, 0, 0, TimeSpan.Zero);
    }

    public sealed record FakeRequest(string Method, string Path, string? App, string? ItemName, string? DriveId, string? ItemId);

    private sealed class Session
    {
        public required string DriveId { get; init; }
        public required string ItemId { get; init; }
        public MemoryStream Data { get; } = new();
        public long Total { get; set; }
    }

    // ---------------------------------------------------------------- setup

    public FakeDrive AddLibrary(string sitePath, string libraryName, string urlName)
    {
        lock (_sync)
        {
            if (!_sites.TryGetValue(sitePath, out var site))
            {
                site = ($"site-{_nextId++}", $"https://{Host}{sitePath}");
                _sites[sitePath] = site;
                _drivesBySite[site.Id] = [];
            }
            string driveId = $"drive-{_nextId++}";
            var root = new FakeItem { Id = $"root-{driveId}", Name = "root", IsFolder = true, IsRoot = true };
            var drive = new FakeDrive { Id = driveId, Name = libraryName, WebUrl = $"{site.WebUrl}/{Uri.EscapeDataString(urlName)}", Root = root };
            drive.Items[root.Id] = root;
            _drives[driveId] = drive;
            _drivesBySite[site.Id].Add(drive);
            return drive;
        }
    }

    // ---------------------------------------------------------------- persistence (for the demo)

    private sealed record SavedSite(string Path, string Id, string WebUrl, List<SavedDrive> Drives);
    private sealed record SavedDrive(string Id, string Name, string WebUrl, List<FakeItem> Items);
    private sealed record Saved(int NextId, int DeltaPageSize, List<SavedSite> Sites);

    /// <summary>Writes the sites, libraries and items (with content and versions) to a JSON file.</summary>
    public void Save(string path)
    {
        Saved saved;
        lock (_sync)
            saved = new Saved(_nextId, DeltaPageSize, [.. _sites.Select(s => new SavedSite(s.Key, s.Value.Id, s.Value.WebUrl,
                [.. _drivesBySite[s.Value.Id].Select(d => new SavedDrive(d.Id, d.Name, d.WebUrl, [.. d.Items.Values]))]))]);
        string temp = path + ".tmp";
        using (var stream = File.Create(temp))
            JsonSerializer.Serialize(stream, saved);
        File.Move(temp, path, overwrite: true);
    }

    public static FakeGraph Load(string path)
    {
        Saved saved;
        using (var stream = File.OpenRead(path))
            saved = JsonSerializer.Deserialize<Saved>(stream) ?? throw new InvalidDataException(path);
        var graph = new FakeGraph { _nextId = saved.NextId, DeltaPageSize = saved.DeltaPageSize };
        foreach (var site in saved.Sites)
        {
            graph._sites[site.Path] = (site.Id, site.WebUrl);
            graph._drivesBySite[site.Id] = [];
            foreach (var d in site.Drives)
            {
                var drive = new FakeDrive { Id = d.Id, Name = d.Name, WebUrl = d.WebUrl, Root = d.Items.Single(i => i.IsRoot) };
                foreach (var item in d.Items)
                    drive.Items[item.Id] = item;
                graph._drives[d.Id] = drive;
                graph._drivesBySite[site.Id].Add(drive);
            }
        }
        return graph;
    }

    /// <summary>Adds a file under a folder path like "Fin/2024", creating folders.</summary>
    public FakeItem AddFile(FakeDrive drive, string path, byte[] content)
    {
        lock (_sync)
        {
            var parts = path.Split('/');
            var parent = drive.Root;
            foreach (string folderName in parts[..^1])
            {
                var folder = drive.Items.Values.FirstOrDefault(i => i.IsFolder && i.ParentId == parent.Id && i.Name == folderName);
                if (folder is null)
                {
                    folder = new FakeItem { Id = $"f{_nextId++}", Name = folderName, ParentId = parent.Id, IsFolder = true };
                    drive.Items[folder.Id] = folder;
                }
                parent = folder;
            }
            var item = new FakeItem { Id = $"i{_nextId++}", Name = parts[^1], ParentId = parent.Id, Content = content };
            drive.Items[item.Id] = item;
            return item;
        }
    }

    public void Remove(FakeDrive drive, FakeItem item)
    {
        lock (_sync)
            drive.Items.Remove(item.Id);
    }

    /// <summary>The next request with an older token of this app gets 401 (token expired).</summary>
    public void ExpireTokens(string app) => _tokenGeneration.AddOrUpdate(app, 1, (_, g) => g + 1);

    public int MinTokenGeneration(string app) => _tokenGeneration.GetValueOrDefault(app);

    // ---------------------------------------------------------------- handling

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        await Task.Yield();
        string url = request.RequestUri!.ToString();
        string? app = null;
        int generation = 0;
        if (request.Headers.Authorization?.Parameter is { } token)
        {
            var parts = token.Split('|');
            app = parts[0];
            generation = int.Parse(parts[1]);
        }
        else if (url.StartsWith(UploadBase, StringComparison.Ordinal))
        {
            app = request.Headers.TryGetValues("X-Test-App", out var v) ? v.First() : null;
        }

        var (driveId, itemId) = ParseIds(url);
        FakeItem? item = null;
        lock (_sync)
        {
            if (driveId is not null && itemId is not null && _drives.TryGetValue(driveId, out var d))
                d.Items.TryGetValue(itemId, out item);
        }
        var fake = new FakeRequest(request.Method.Method, url, app, item?.Name, driveId, itemId);
        Log.Enqueue(fake);

        foreach (var fault in Faults.ToArray())
        {
            if (fault(fake) is { } faulted)
                return faulted;
        }

        if (!url.StartsWith(UploadBase, StringComparison.Ordinal))
        {
            if (!request.Headers.UserAgent.ToString().StartsWith("NONISV|", StringComparison.Ordinal))
                return Error(HttpStatusCode.BadRequest, "badUserAgent");
            if (app is null)
                return Error(HttpStatusCode.Unauthorized, "InvalidAuthenticationToken");
            if (generation < MinTokenGeneration(app))
                return Error(HttpStatusCode.Unauthorized, "InvalidAuthenticationToken");
        }

        try
        {
            return await Route(request, url, fake, ct);
        }
        catch (KeyNotFoundException)
        {
            return Error(HttpStatusCode.NotFound, "itemNotFound");
        }
    }

    private static (string? Drive, string? Item) ParseIds(string url)
    {
        var path = new Uri(url).AbsolutePath;
        var segs = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
        string? drive = null, item = null;
        for (int i = 0; i < segs.Length - 1; i++)
        {
            if (segs[i] == "drives")
                drive = segs[i + 1];
            if (segs[i] == "items")
                item = segs[i + 1];
        }
        return (drive, item);
    }

    private async Task<HttpResponseMessage> Route(HttpRequestMessage request, string url, FakeRequest fake, CancellationToken ct)
    {
        var uri = new Uri(url);
        string path = Uri.UnescapeDataString(uri.AbsolutePath);
        string method = request.Method.Method;

        if (url.StartsWith(UploadBase, StringComparison.Ordinal))
            return await HandleSession(request, url[UploadBase.Length..], fake, ct);

        if (!path.StartsWith("/v1.0/", StringComparison.Ordinal))
            return Error(HttpStatusCode.NotFound, "badPath");
        path = path["/v1.0".Length..];

        lock (_sync)
        {
            if (path == "/sites" && uri.Query.Contains("search=", StringComparison.Ordinal))
            {
                string q = Uri.UnescapeDataString(uri.Query.Split("search=")[1].Split('&')[0]);
                var found = new JsonArray();
                foreach (var (sitePath, site) in _sites.Where(s => s.Key.Contains(q, StringComparison.OrdinalIgnoreCase)))
                    found.Add(new JsonObject { ["id"] = site.Id, ["webUrl"] = site.WebUrl, ["displayName"] = sitePath });
                return Json(new JsonObject { ["value"] = found });
            }
            if (path.EndsWith("/children", StringComparison.Ordinal))
            {
                var cs = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
                var d = _drives[cs[1]];
                string parentId = cs[2] == "root" ? d.Root.Id : cs[3];
                var children = new JsonArray();
                foreach (var child in d.Items.Values.Where(i => i.ParentId == parentId))
                    children.Add(ItemJson(d, child));
                return Json(new JsonObject { ["value"] = children });
            }
            // /sites/{host}:/{path}  or /sites/{host}
            if (path.StartsWith("/sites/", StringComparison.Ordinal) && !path.Contains("/drives", StringComparison.Ordinal))
            {
                string rest = path["/sites/".Length..];
                string sitePath = rest.Contains(":/") ? "/" + rest[(rest.IndexOf(":/", StringComparison.Ordinal) + 2)..].Trim('/') : "/";
                if (!_sites.TryGetValue(sitePath, out var site))
                    return Error(HttpStatusCode.NotFound, "itemNotFound");
                return Json(new JsonObject { ["id"] = site.Id, ["webUrl"] = site.WebUrl });
            }
            if (path.StartsWith("/sites/", StringComparison.Ordinal) && path.EndsWith("/drives", StringComparison.Ordinal))
            {
                string siteId = path.Split('/')[2];
                var arr = new JsonArray();
                foreach (var d in _drivesBySite[siteId])
                    arr.Add(new JsonObject { ["id"] = d.Id, ["name"] = d.Name, ["webUrl"] = d.WebUrl });
                return Json(new JsonObject { ["value"] = arr });
            }

            var segs = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
            var drive = _drives[segs[1]];
            if (path.Contains("/root:/", StringComparison.Ordinal))
            {
                string itemPath = path[(path.IndexOf("/root:/", StringComparison.Ordinal) + 7)..].Trim('/');
                var found = FindByPath(drive, itemPath) ?? throw new KeyNotFoundException();
                return Json(ItemJson(drive, found));
            }
            if (path.EndsWith("/root/delta", StringComparison.Ordinal))
                return Delta(drive, uri);

            var item = drive.Items[segs[3]];
            if (segs.Length == 4 && method == "GET")
                return Json(ItemJson(drive, item, full: true));
            if (segs.Length == 7 && segs[4] == "versions" && segs[6] == "content" && method == "GET")
            {
                int v = int.Parse(segs[5].Split('.')[0], System.Globalization.CultureInfo.InvariantCulture);
                byte[] content = v == item.Version ? item.Content : item.History.TryGetValue(v, out var old) ? old : throw new KeyNotFoundException();
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(content.ToArray()) };
            }
            if (segs.Length == 5 && segs[4] == "content" && method == "GET")
            {
                OnDownloaded?.Invoke(fake, item);
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(item.Content.ToArray()) };
            }
            if (segs.Length == 5 && segs[4] == "content" && method == "PUT")
            {
                if (Precondition(request, item) is { } failed)
                    return failed;
                byte[] body = request.Content!.ReadAsByteArrayAsync(ct).GetAwaiter().GetResult();
                item.History[item.Version] = item.Content;
                item.Content = body;
                item.Version++;
                item.Uploads++;
                item.Modified = DateTimeOffset.UtcNow;
                OnUploaded?.Invoke(fake, item);
                return Json(ItemJson(drive, item));
            }
            if (segs.Length == 5 && segs[4] == "createUploadSession" && method == "POST")
            {
                if (Precondition(request, item) is { } failed)
                    return failed;
                string id = Guid.NewGuid().ToString("N");
                _sessions[id] = new Session { DriveId = drive.Id, ItemId = item.Id };
                return Json(new JsonObject { ["uploadUrl"] = UploadBase + id });
            }
        }
        return Error(HttpStatusCode.BadRequest, "notImplemented");
    }

    private HttpResponseMessage? Precondition(HttpRequestMessage request, FakeItem item)
    {
        if (item.LockedUploads > 0)
        {
            item.LockedUploads--;
            return Error((HttpStatusCode)423, "resourceLocked");
        }
        if (item.CheckedOutBy is not null)
            return Error((HttpStatusCode)423, "resourceLocked");
        string? ifMatch = request.Headers.TryGetValues("If-Match", out var v) ? v.First() : null;
        if (ifMatch is null)
            return Error(HttpStatusCode.BadRequest, "ifMatchRequiredInTest");
        if (ifMatch != item.ETag)
            return Error(HttpStatusCode.PreconditionFailed, "resourceModified");
        return null;
    }

    private async Task<HttpResponseMessage> HandleSession(HttpRequestMessage request, string id, FakeRequest fake, CancellationToken ct)
    {
        Session session;
        lock (_sync)
        {
            if (!_sessions.TryGetValue(id, out session!))
                return Error(HttpStatusCode.NotFound, "itemNotFound");
        }
        if (request.Method == HttpMethod.Get)
            return Json(new JsonObject { ["nextExpectedRanges"] = new JsonArray($"{session.Data.Length}-") });
        if (request.Method == HttpMethod.Delete)
        {
            lock (_sync)
                _sessions.Remove(id);
            return new HttpResponseMessage(HttpStatusCode.NoContent);
        }

        var range = request.Content!.Headers.ContentRange!;
        byte[] body = await request.Content.ReadAsByteArrayAsync(ct);
        lock (_sync)
        {
            if (range.From != session.Data.Length)
                return Error((HttpStatusCode)416, "invalidRange");
            session.Data.Write(body);
            session.Total = range.Length!.Value;
            if (session.Data.Length < session.Total)
            {
                var accepted = new HttpResponseMessage(HttpStatusCode.Accepted)
                {
                    Content = new StringContent(new JsonObject { ["nextExpectedRanges"] = new JsonArray($"{session.Data.Length}-") }.ToJsonString(), Encoding.UTF8, "application/json"),
                };
                return accepted;
            }
            var drive = _drives[session.DriveId];
            var item = drive.Items[session.ItemId];
            item.History[item.Version] = item.Content;
            item.Content = session.Data.ToArray();
            item.Version++;
            item.Uploads++;
            _sessions.Remove(id);
            OnUploaded?.Invoke(fake, item);
            var done = Json(ItemJson(drive, item));
            done.StatusCode = HttpStatusCode.Created;
            return done;
        }
    }

    private HttpResponseMessage Delta(FakeDrive drive, Uri uri)
    {
        var all = drive.Items.Values.OrderBy(i => i.IsFile() ? 1 : 0).ThenBy(i => int.Parse(new string(i.Id.Where(char.IsDigit).DefaultIfEmpty('0').ToArray())))
            .ToList();
        if (FilesBeforeFolders)
            all = all.OrderBy(i => i.IsRoot ? 0 : i.IsFolder ? 2 : 1).ToList();
        int start = uri.Query.Contains("token=", StringComparison.Ordinal)
            ? int.Parse(uri.Query.Split("token=")[1].Split('&')[0])
            : 0;
        var page = all.Skip(start).Take(DeltaPageSize).ToList();
        var value = new JsonArray();
        foreach (var i in page)
            value.Add(ItemJson(drive, i));
        var obj = new JsonObject { ["value"] = value };
        if (start + DeltaPageSize < all.Count)
            obj["@odata.nextLink"] = $"{Base}/drives/{drive.Id}/root/delta?token={start + DeltaPageSize}";
        else
            obj["@odata.deltaLink"] = $"{Base}/drives/{drive.Id}/root/delta?token=latest";
        return Json(obj);
    }

    private static FakeItem? FindByPath(FakeDrive drive, string path)
    {
        var current = drive.Root;
        foreach (string part in path.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            current = drive.Items.Values.FirstOrDefault(i => i.ParentId == current.Id && string.Equals(i.Name, part, StringComparison.OrdinalIgnoreCase));
            if (current is null)
                return null;
        }
        return current;
    }

    private static JsonObject ItemJson(FakeDrive drive, FakeItem i, bool full = false)
    {
        var o = new JsonObject
        {
            ["id"] = i.Id,
            ["name"] = i.Name,
            ["eTag"] = i.ETag,
            ["cTag"] = i.ETag,
            ["size"] = i.Content.Length,
            ["lastModifiedDateTime"] = i.Modified.ToString("o"),
            ["lastModifiedBy"] = new JsonObject { ["user"] = new JsonObject { ["email"] = "user@contoso.com" } },
        };
        if (i.ParentId is not null)
            o["parentReference"] = new JsonObject { ["id"] = i.ParentId, ["driveId"] = drive.Id };
        if (i.IsRoot)
            o["root"] = new JsonObject();
        if (i.IsFolder)
            o["folder"] = new JsonObject { ["childCount"] = 0 };
        else
            o["file"] = new JsonObject { ["mimeType"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
        if (full && !i.IsFolder)
        {
            o["publication"] = new JsonObject { ["level"] = i.CheckedOutBy is null ? "published" : "checkout", ["versionId"] = $"{i.Version}.0" };
            o["listItem"] = new JsonObject
            {
                ["id"] = "1",
                ["fields"] = new JsonObject { ["_UIVersionString"] = $"{i.Version}.0", ["CheckoutUser"] = i.CheckedOutBy is null ? null : new JsonObject { ["LookupValue"] = i.CheckedOutBy } },
            };
        }
        return o;
    }

    public static HttpResponseMessage Json(JsonNode node) => new(HttpStatusCode.OK)
    {
        Content = new StringContent(node.ToJsonString(), Encoding.UTF8, "application/json"),
    };

    public static HttpResponseMessage Error(HttpStatusCode status, string code, int? retryAfterSeconds = null)
    {
        var response = new HttpResponseMessage(status)
        {
            Content = new StringContent(JsonSerializer.Serialize(new { error = new { code, message = "fake " + code } }), Encoding.UTF8, "application/json"),
        };
        response.Headers.TryAddWithoutValidation("request-id", Guid.NewGuid().ToString());
        if (retryAfterSeconds is int s)
            response.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.FromSeconds(s));
        return response;
    }

    public int UploadsOf(FakeItem item)
    {
        lock (_sync)
            return item.Uploads;
    }

    /// <summary>RU spent by every app according to the requests the fake has seen (same cost table as the client).</summary>
    public int CostOf(FakeRequest r) => r.Method switch
    {
        "GET" when r.Path.Contains("/content", StringComparison.Ordinal) => RequestCost.Download,
        "GET" when r.Path.Contains("/delta", StringComparison.Ordinal) => r.Path.Contains("token=", StringComparison.Ordinal) ? RequestCost.DeltaWithToken : RequestCost.List,
        "GET" when r.Path.EndsWith("/drives?$select=id,name,webUrl", StringComparison.Ordinal) => RequestCost.List,
        "GET" => RequestCost.Get,
        "PUT" when r.Path.StartsWith(UploadBase, StringComparison.Ordinal) => RequestCost.UploadChunk,
        _ => RequestCost.Upload,
    };
}

internal static class FakeItemExtensions
{
    public static bool IsFile(this FakeGraph.FakeItem i) => !i.IsFolder;
}

/// <summary>Token "app|generation"; a forced refresh moves to the current generation of the fake.</summary>
public sealed class FakeTokens(FakeGraph graph, string app) : ITokenProvider
{
    private int _generation;

    public int Refreshes;

    public Task<string> GetTokenAsync(bool forceRefresh, CancellationToken ct)
    {
        if (forceRefresh)
        {
            Interlocked.Increment(ref Refreshes);
            _generation = graph.MinTokenGeneration(app);
        }
        return Task.FromResult($"{app}|{_generation}");
    }
}

public sealed class FakeAppFactory(FakeGraph graph) : ISpoAppFactory
{
    public ConcurrentDictionary<string, SpoApp> Created { get; } = new();

    public ConcurrentDictionary<string, FakeTokens> Tokens { get; } = new();

    public int ChunkSize { get; set; } = 320 * 1024;

    public SpoApp Create(Connection connection, SpoAppSettings settings, RuBudget tenantBudget)
    {
        var (perMinute, perDay, _) = connection.Limits();
        var budget = new RuBudget(connection.Name, perMinute, TimeSpan.FromMinutes(1), perDay) { Scale = settings.BudgetScale };
        var throttle = new ThrottleState(connection.Name, settings.MaxConcurrency, settings.BlockedAfter);
        var tokens = Tokens.GetOrAdd(connection.Name, n => new FakeTokens(graph, n));
        var http = new HttpClient(new AppHeaderHandler(connection.Name) { InnerHandler = graph }, disposeHandler: false) { Timeout = Timeout.InfiniteTimeSpan };
        var client = new GraphClient(connection.Name, http, tokens, budget, tenantBudget, throttle, new GraphClientOptions
        {
            BaseUrl = FakeGraph.Base,
            RequestAttempts = settings.RequestAttempts,
            RequestTimeout = settings.RequestTimeout,
            BaseBackoff = TimeSpan.FromMilliseconds(10),
            MaxBackoff = TimeSpan.FromMilliseconds(100),
            ChunkSize = ChunkSize,
        }, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance);
        var app = new SpoApp(connection.Name, client, budget, throttle);
        Created[connection.Name] = app;
        return app;
    }

    /// <summary>Marks unauthenticated upload-session requests with the app name, for attribution in the fake.</summary>
    private sealed class AppHeaderHandler(string app) : DelegatingHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            request.Headers.TryAddWithoutValidation("X-Test-App", app);
            return base.SendAsync(request, cancellationToken);
        }
    }
}

/// <summary>Minimal xlsx packages for engine tests.</summary>
public static class TestXlsx
{
    public static byte[] WithHyperlink(string url, string sharedText = "text", int padBytes = 0)
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string text)
            {
                using var s = zip.CreateEntry(name).Open();
                s.Write(Encoding.UTF8.GetBytes(text));
            }
            Add("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
            Add("xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
            Add("xl/worksheets/sheet1.xml", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>");
            Add("xl/worksheets/_rels/sheet1.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                $"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{url}\" TargetMode=\"External\"/>" +
                "</Relationships>");
            Add("xl/sharedStrings.xml", $"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><si><t>{sharedText}</t></si></sst>");
            if (padBytes > 0)
            {
                var pad = zip.CreateEntry("xl/media/pad.bin", CompressionLevel.NoCompression);
                using var s = pad.Open();
                var bytes = new byte[padBytes];
                new Random(42).NextBytes(bytes);
                s.Write(bytes);
            }
        }
        return ms.ToArray();
    }

    public static string HyperlinkOf(byte[] xlsx)
    {
        using var zip = new ZipArchive(new MemoryStream(xlsx));
        using var reader = new StreamReader(zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels")!.Open());
        string rels = reader.ReadToEnd();
        int i = rels.IndexOf("Target=\"", StringComparison.Ordinal) + 8;
        return rels[i..rels.IndexOf('"', i)];
    }
}
