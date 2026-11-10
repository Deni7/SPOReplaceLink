using System.Text.Json;

namespace SPOReplaceLink.Spo.Graph;

public sealed record SpoSite(string Id, string WebUrl);

public sealed record SpoDrive(string Id, string Name, string WebUrl);

/// <summary>Drive item with what the run needs (ENU-3).</summary>
public sealed record SpoItem(
    string Id,
    string Name,
    string? ParentId,
    bool IsFolder,
    bool IsFile,
    bool IsRoot,
    bool Deleted,
    long Size,
    string? ETag,
    string? CTag,
    DateTimeOffset? Modified,
    string? ModifiedBy,
    string? Version = null,
    string? CheckedOutBy = null)
{
    public static SpoItem Parse(JsonElement e)
    {
        string? Str(JsonElement el, string name) => el.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

        string? parent = e.TryGetProperty("parentReference", out var p) ? Str(p, "id") : null;
        string? modifiedBy = e.TryGetProperty("lastModifiedBy", out var mb) && mb.TryGetProperty("user", out var u)
            ? Str(u, "email") ?? Str(u, "displayName")
            : null;
        DateTimeOffset? modified = Str(e, "lastModifiedDateTime") is { } m && DateTimeOffset.TryParse(m, out var dt) ? dt : null;

        string? version = null;
        string? checkedOut = null;
        if (e.TryGetProperty("publication", out var pub) && pub.ValueKind == JsonValueKind.Object)
        {
            version = Str(pub, "versionId");
            if (Str(pub, "level") == "checkout")
                checkedOut = "unknown";
        }
        if (e.TryGetProperty("listItem", out var li) && li.TryGetProperty("fields", out var f))
        {
            version ??= Str(f, "_UIVersionString");
            if (f.TryGetProperty("CheckoutUser", out var cu) && cu.ValueKind != JsonValueKind.Null)
                checkedOut = cu.ValueKind == JsonValueKind.Object ? Str(cu, "LookupValue") ?? "unknown" : cu.ToString();
        }

        return new SpoItem(
            Str(e, "id") ?? string.Empty,
            Str(e, "name") ?? string.Empty,
            parent,
            e.TryGetProperty("folder", out _),
            e.TryGetProperty("file", out _),
            e.TryGetProperty("root", out _),
            e.TryGetProperty("deleted", out _),
            e.TryGetProperty("size", out var s) && s.ValueKind == JsonValueKind.Number ? s.GetInt64() : 0,
            Str(e, "eTag"),
            Str(e, "cTag"),
            modified,
            modifiedBy,
            version,
            checkedOut);
    }
}

public sealed record DeltaPage(IReadOnlyList<SpoItem> Items, string? NextLink, string? DeltaLink);

/// <summary>THR-1: cost of requests in resource units, from the Microsoft table.</summary>
public static class RequestCost
{
    public const int Get = 1;
    public const int DeltaWithToken = 1;
    public const int List = 2;
    public const int Download = 1;
    public const int Upload = 2;
    public const int UploadChunk = 1;
}
