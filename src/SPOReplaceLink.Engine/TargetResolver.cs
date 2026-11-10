using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Spo;
using SPOReplaceLink.Spo.Graph;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Engine;

public static class TargetResolver
{
    /// <summary>Site → library → folder of a target (GUI-17 by URL: works with Sites.Selected).</summary>
    public static async Task<(SpoDrive Drive, string? RootItem, string RootPath)> ResolveAsync(GraphClient client, Target target, CancellationToken ct)
    {
        var site = await client.GetSiteAsync(target.SiteUrl!, ct).ConfigureAwait(false);
        var drives = await client.GetDrivesAsync(site.Id, ct).ConfigureAwait(false);
        string wanted = target.Library!.Trim().TrimEnd('/');
        var drive = drives.FirstOrDefault(d =>
                string.Equals(d.Name, wanted, StringComparison.OrdinalIgnoreCase)
                || string.Equals(Uri.UnescapeDataString(d.WebUrl.TrimEnd('/')), Uri.UnescapeDataString(wanted), StringComparison.OrdinalIgnoreCase)
                || string.Equals(Uri.UnescapeDataString(d.WebUrl.TrimEnd('/').Split('/')[^1]), wanted, StringComparison.OrdinalIgnoreCase))
            ?? throw new SpoException(ErrorClass.File, L.F("бібліотеку «{0}» не знайдено на сайті {1} (є: {2}).", target.Library, target.SiteUrl, string.Join(", ", drives.Select(d => d.Name))));

        string folder = (target.Folder ?? string.Empty).Replace('\\', '/').Trim('/');
        if (folder.Length == 0)
            return (drive, null, string.Empty);
        var item = await client.GetItemByPathAsync(drive.Id, folder, ct).ConfigureAwait(false);
        if (!item.IsFolder)
            throw new SpoException(ErrorClass.File, L.F("«{0}» у бібліотеці {1} — не папка.", folder, drive.Name));
        return (drive, item.Id, folder);
    }

}
