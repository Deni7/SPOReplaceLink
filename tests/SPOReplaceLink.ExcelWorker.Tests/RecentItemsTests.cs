using System.Runtime.Versioning;

namespace SPOReplaceLink.ExcelWorker.Tests;

/// <summary>Shortcuts to Excel's temporary copies leave the Recent folder; the user's own shortcuts stay.</summary>
public sealed class RecentItemsTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "spo-rl-recent", Guid.NewGuid().ToString("N"));
    private readonly string _root;
    private readonly string _recent;

    public RecentItemsTests()
    {
        _root = Directory.CreateDirectory(Path.Combine(_dir, "temp")).FullName;
        _recent = Directory.CreateDirectory(Path.Combine(_dir, "Recent")).FullName;
    }

    [SupportedOSPlatform("windows")]
    private string Link(string name, string target)
    {
        string lnk = Path.Combine(_recent, name + ".lnk");
        RecentItems.Create(lnk, target);
        return lnk;
    }

    [Fact]
    public void RemovesOnlyShortcutsIntoTheTemporaryFolder()
    {
        if (!OperatingSystem.IsWindows())
            return;
        string ours = Link("book.xls", Path.Combine(_root, "run-2", "x66830-2d4511", "1", "2", "book.xls"));
        string oursOut = Link("book-out.xls", Path.Combine(_root, "run-2", "x66831-092a63", "book-out.xls"));
        string mine = Link("Report.xlsx", Path.Combine(_dir, "Documents", "Report.xlsx"));
        string similar = Link("temp2.xlsx", _root + "2\\book.xls"); // a sibling folder whose name starts the same

        Assert.Equal(Path.Combine(_root, "run-2", "x66831-092a63", "book-out.xls"), RecentItems.Target(oursOut), ignoreCase: true);
        Assert.Equal(2, new RecentItems(_root, _recent).Sweep());
        Assert.False(File.Exists(ours));
        Assert.False(File.Exists(oursOut));
        Assert.True(File.Exists(mine));
        Assert.True(File.Exists(similar));
    }

    [Fact]
    public void LaterSweepsLookOnlyAtNewShortcuts()
    {
        if (!OperatingSystem.IsWindows())
            return;
        var recent = new RecentItems(_root, _recent);
        Assert.Equal(0, recent.Sweep());
        // Older than the previous sweep: not looked at again (the first sweep of the next start finds it).
        string old = Link("old.xls", Path.Combine(_root, "old.xls"));
        File.SetLastWriteTimeUtc(old, DateTime.UtcNow.AddHours(-1));
        string fresh = Link("book.xlsb", Path.Combine(_root, "x1", "book.xlsb"));

        Assert.Equal(1, recent.Sweep());
        Assert.True(File.Exists(old));
        Assert.False(File.Exists(fresh));
        Assert.Equal(1, new RecentItems(_root, _recent).Sweep());
        Assert.False(File.Exists(old));
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_dir, true);
        }
        catch (IOException)
        {
        }
    }
}
