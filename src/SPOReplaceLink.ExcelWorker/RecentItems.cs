using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
using System.Text;

namespace SPOReplaceLink.ExcelWorker;

/// <summary>
/// Some Excel versions add every workbook they open or save to the Windows Recent Items, and Explorer then lists the
/// temporary copies (book.xls, book-out.xls, ...) in Quick Access instead of the user's own files. This removes the
/// shortcuts in the Recent folder that point under the temporary folder; other shortcuts and Windows settings stay as they are.
/// </summary>
internal sealed class RecentItems
{
    private readonly string _recent;
    private readonly string[] _roots;
    private readonly object _sync = new();
    private DateTime _since = DateTime.MinValue;

    /// <param name="root">The temporary folder whose files must not stay in Recent Items.</param>
    /// <param name="recentFolder">The Recent folder; the user's one by default (tests pass their own).</param>
    public RecentItems(string root, string? recentFolder = null)
    {
        _recent = recentFolder ?? (OperatingSystem.IsWindows() ? Environment.GetFolderPath(Environment.SpecialFolder.Recent) : string.Empty);
        string full = Path.GetFullPath(root).TrimEnd('\\') + "\\";
        // A shortcut may hold the 8.3 form of the path (C:\Users\LONGNA~1\...).
        _roots = ShortPath(full) is { } shortForm && !shortForm.Equals(full, StringComparison.OrdinalIgnoreCase) ? [full, shortForm] : [full];
    }

    /// <summary>
    /// Deletes shortcuts to files under the root that appeared since the previous call; the first call looks at all of them
    /// (left by an earlier run that stopped before it could clean up). Returns how many were deleted.
    /// </summary>
    public int Sweep()
    {
        if (!OperatingSystem.IsWindows())
            return 0;
        if (_recent.Length == 0 || !Monitor.TryEnter(_sync))
            return 0; // another thread is sweeping
        try
        {
            // Windows writes the shortcut a moment after Excel opens the file: look a little further back.
            var since = _since == DateTime.MinValue ? DateTime.MinValue : _since - TimeSpan.FromMinutes(1);
            _since = DateTime.UtcNow;
            int deleted = 0;
            foreach (var file in new DirectoryInfo(_recent).EnumerateFiles("*.lnk"))
            {
                try
                {
                    if (file.LastWriteTimeUtc < since || !IsUnderRoot(Target(file.FullName)))
                        continue;
                    file.Delete();
                    deleted++;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or COMException)
                {
                    // In use or not a readable shortcut: left alone.
                }
            }
            return deleted;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return 0; // no Recent folder
        }
        finally
        {
            Monitor.Exit(_sync);
        }
    }

    private bool IsUnderRoot(string? target) =>
        target is not null && _roots.Any(r => target.StartsWith(r, StringComparison.OrdinalIgnoreCase));

    /// <summary>The path a shortcut points to, read from the shortcut itself (the target is usually deleted by now).</summary>
    [SupportedOSPlatform("windows")]
    internal static string? Target(string lnk)
    {
        object link = new ShellLink();
        try
        {
            ((IPersistFile)link).Load(lnk, 0); // STGM_READ
            var path = new StringBuilder(1024);
            ((IShellLinkW)link).GetPath(path, path.Capacity, IntPtr.Zero, 0);
            return path.Length > 0 ? path.ToString() : null;
        }
        finally
        {
            Marshal.FinalReleaseComObject(link);
        }
    }

    /// <summary>Test aid: a shortcut to <paramref name="target"/> (which need not exist).</summary>
    [SupportedOSPlatform("windows")]
    internal static void Create(string lnk, string target)
    {
        object link = new ShellLink();
        try
        {
            ((IShellLinkW)link).SetPath(target);
            ((IPersistFile)link).Save(lnk, true);
        }
        finally
        {
            Marshal.FinalReleaseComObject(link);
        }
    }

    private static string? ShortPath(string path)
    {
        if (!OperatingSystem.IsWindows())
            return null;
        var buffer = new StringBuilder(1024);
        return GetShortPathNameW(path, buffer, buffer.Capacity) is > 0 and < 1024 ? buffer.ToString() : null;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetShortPathNameW(string longPath, StringBuilder shortPath, int size);

    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    private class ShellLink;

    [ComImport]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellLinkW
    {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder file, int size, IntPtr findData, int flags);
        void GetIDList(out IntPtr idList);
        void SetIDList(IntPtr idList);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder name, int size);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string name);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder dir, int size);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string dir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder args, int size);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string args);
        void GetHotkey(out short hotkey);
        void SetHotkey(short hotkey);
        void GetShowCmd(out int showCmd);
        void SetShowCmd(int showCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder path, int size, out int icon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string path, int icon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string path, int reserved);
        void Resolve(IntPtr hwnd, int flags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string file);
    }
}
