using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Urls;

namespace SPOReplaceLink.OpenXml;

/// <summary>
/// The workbook is an intermediate copy that Excel saved in <see cref="Folder"/> (ExcelWorker converts xls and xlsb
/// to xlsm and back). On save Excel writes links to files on the same drive relative to that folder, so a link that
/// was absolute (C:\Old\a.xlsx) comes out as ../../../Old/a.xlsx, /Old/a.xlsx or file:///C:\Old\a.xlsx.
/// <see cref="Folder"/> lies several levels below <see cref="Root"/>: a location inside Root is what a link relative to
/// the document became (it is resolved against the old document folder as usual), a location outside Root was an
/// absolute local path (it is matched against the rules as that path).
/// </summary>
public sealed record LocalCopy(string Folder, string Root)
{
    /// <summary>How many levels <see cref="Folder"/> lies below <see cref="Root"/>; deeper relative links are not expected.</summary>
    public const int Depth = 8;

    /// <summary>Makes the folder layout for one file under <paramref name="root"/> and returns it.</summary>
    public static LocalCopy Create(string root)
    {
        string folder = root;
        for (int i = 1; i <= Depth; i++)
            folder = Path.Combine(folder, i.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Directory.CreateDirectory(folder);
        return new LocalCopy(Path.GetFullPath(folder), Path.GetFullPath(root));
    }

    /// <summary>The link as it was before Excel rewrote it, or null if it stays as written.</summary>
    public string? Translate(string value, LinkKind kind)
    {
        if (kind is not (LinkKind.LinkSource or LinkKind.Hyperlink or LinkKind.OtherExternal))
            return null;
        try
        {
            switch (LinkClassifier.Classify(value))
            {
                case LinkForm.Relative:
                {
                    string path = Path.GetFullPath(Path.Combine(Folder, Unescape(value)));
                    return IsInside(path) ? null : "file:///" + path;
                }
                case LinkForm.ServerRelative when kind == LinkKind.LinkSource:
                {
                    // "/Old/a.xlsx": Excel's form for a file on the same drive as the workbook.
                    string drive = Path.GetPathRoot(Folder)!.TrimEnd('\\', '/');
                    string path = Path.GetFullPath(drive + Unescape(value).Replace('/', '\\'));
                    return IsInside(path) ? RelativeOf(path) : "file:///" + path;
                }
                case LinkForm.File or LinkForm.Drive:
                {
                    string? path = LocalPath(value);
                    return path is not null && IsInside(path) ? RelativeOf(path) : null;
                }
                default:
                    return null;
            }
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return null;
        }
    }

    private bool IsInside(string path) =>
        path.StartsWith(Root.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase);

    private string RelativeOf(string path) => Path.GetRelativePath(Folder, path).Replace('\\', '/');

    private static string? LocalPath(string value)
    {
        string s = value;
        if (s.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
            s = s[5..].TrimStart('/', '\\');
        s = Unescape(s);
        return LinkClassifier.Classify(s) == LinkForm.Drive ? Path.GetFullPath(s) : null;
    }

    private static string Unescape(string s)
    {
        try
        {
            return Uri.UnescapeDataString(s);
        }
        catch (UriFormatException)
        {
            return s;
        }
    }
}
