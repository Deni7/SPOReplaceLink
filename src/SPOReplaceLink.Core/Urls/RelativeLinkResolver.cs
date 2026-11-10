namespace SPOReplaceLink.Core.Urls;

/// <summary>
/// Resolves a relative link (../a.xlsx, Folder\b.xlsx) against the folder the document lived in
/// before migration (RUL-6). Works on text segments so the original spelling of each segment is kept.
/// </summary>
public static class RelativeLinkResolver
{
    /// <returns>The absolute link, or null if it climbs above the root or the base is not absolute.</returns>
    public static string? Resolve(string documentFolder, string relative)
    {
        if (!TrySplitRoot(documentFolder, out string root, out char separator, out string path))
            return null;

        var segments = Split(path);
        foreach (string segment in Split(relative))
        {
            if (segment == ".")
                continue;
            if (segment == "..")
            {
                if (segments.Count == 0)
                    return null;
                segments.RemoveAt(segments.Count - 1);
                continue;
            }
            segments.Add(segment);
        }

        string joined = string.Join(separator, segments);
        if (relative.Length > 0 && LinkClassifier.IsSlash(relative[^1]))
            joined += separator;
        return root + separator + joined;
    }

    /// <summary>https://host, \\server or C: — what a path starting with a slash is relative to.</summary>
    public static string? RootOf(string documentFolder) =>
        TrySplitRoot(documentFolder, out string root, out _, out _) ? root : null;

    private static List<string> Split(string path) =>
        path.Split('/', '\\').Where(s => s.Length > 0).ToList();

    private static bool TrySplitRoot(string folder, out string root, out char separator, out string path)
    {
        root = path = string.Empty;
        separator = '/';
        switch (LinkClassifier.Classify(folder))
        {
            case LinkForm.Http:
            {
                int hostStart = folder.IndexOf("://", StringComparison.Ordinal) + 3;
                int pathStart = folder.IndexOf('/', hostStart);
                root = pathStart < 0 ? folder : folder[..pathStart];
                path = pathStart < 0 ? string.Empty : folder[pathStart..];
                return true;
            }
            case LinkForm.Unc:
            {
                int hostEnd = folder.IndexOfAny(['/', '\\'], 2);
                root = hostEnd < 0 ? folder : folder[..hostEnd];
                path = hostEnd < 0 ? string.Empty : folder[hostEnd..];
                separator = folder[0];
                return true;
            }
            case LinkForm.Drive:
                root = folder[..2];
                path = folder[2..];
                separator = folder[2];
                return true;
            default:
                return false;
        }
    }
}
