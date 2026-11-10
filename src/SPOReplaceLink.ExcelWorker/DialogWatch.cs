using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;

namespace SPOReplaceLink.ExcelWorker;

/// <summary>
/// Excel may still ask a question although alerts are off and a password is passed: some versions show the password
/// prompt for some files. Nobody answers it, so Excel would sit on the desktop until the timeout and then be killed.
/// While a request runs, every dialog of our Excel is closed (as Cancel) at once; a dialog that does not go away within
/// a few seconds gets the Excel process killed. The question is remembered, so the file is reported for what it asked.
/// </summary>
[SupportedOSPlatform("windows")]
internal sealed class DialogWatch : IDisposable
{
    private const uint WmClose = 0x0010;
    private static readonly TimeSpan KillAfter = TimeSpan.FromSeconds(5);

    private readonly int _pid;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new();
    private volatile bool _active;
    private volatile string? _question;
    private DateTime? _since;

    public DialogWatch(int pid)
    {
        _pid = pid;
        _thread = new Thread(Loop) { IsBackground = true, Name = "Excel dialogs" };
        _thread.Start();
    }

    /// <summary>The title and text of the first dialog Excel showed during the current request, or null.</summary>
    public string? Question => _question;

    /// <summary>Excel did not close a dialog and was killed.</summary>
    public bool Killed { get; private set; }

    public void Begin()
    {
        _question = null;
        _since = null;
        _active = true;
    }

    /// <returns>The question Excel asked during the request, or null.</returns>
    public string? End()
    {
        _active = false;
        return _question;
    }

    private void Loop()
    {
        while (!_stop.Wait(200))
        {
            if (!_active || Killed)
                continue;
            var dialogs = Dialogs(_pid);
            if (dialogs.Count == 0)
            {
                _since = null;
                continue;
            }
            _question ??= dialogs[0].Text;
            _since ??= DateTime.UtcNow;
            if (DateTime.UtcNow - _since > KillAfter)
            {
                Kill();
                continue;
            }
            foreach (var (hwnd, _) in dialogs)
                PostMessage(hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
        }
    }

    private void Kill()
    {
        try
        {
            using var excel = Process.GetProcessById(_pid);
            excel.Kill();
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
        }
        Killed = true;
    }

    /// <summary>Visible dialog windows of the process: Office's own (bosa_sdm, NUIDialog) and standard ones (#32770).</summary>
    internal static List<(IntPtr Hwnd, string Text)> Dialogs(int pid)
    {
        var found = new List<(IntPtr, string)>();
        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out uint owner);
            if (owner != pid || !IsWindowVisible(hwnd))
                return true;
            string cls = ClassName(hwnd);
            if (cls == "#32770" || cls == "NUIDialog" || cls.StartsWith("bosa_sdm", StringComparison.Ordinal))
                found.Add((hwnd, Describe(hwnd)));
            return true;
        }, IntPtr.Zero);
        return found;
    }

    /// <summary>The title, and for standard dialogs the static texts in it ("'book.xls' is protected", ...).</summary>
    private static string Describe(IntPtr hwnd)
    {
        var parts = new List<string> { WindowText(hwnd) };
        EnumChildWindows(hwnd, (child, _) =>
        {
            if (ClassName(child) == "Static" && WindowText(child) is { Length: > 0 } text)
                parts.Add(text);
            return true;
        }, IntPtr.Zero);
        return string.Join(": ", parts.Where(p => p.Length > 0).Select(p => p.Replace('\n', ' ').Trim()));
    }

    private static string ClassName(IntPtr hwnd)
    {
        var sb = new StringBuilder(256);
        return GetClassName(hwnd, sb, sb.Capacity) > 0 ? sb.ToString() : string.Empty;
    }

    private static string WindowText(IntPtr hwnd)
    {
        var sb = new StringBuilder(512);
        return GetWindowText(hwnd, sb, sb.Capacity) > 0 ? sb.ToString() : string.Empty;
    }

    public void Dispose()
    {
        _stop.Set();
        _thread.Join(2_000);
        _stop.Dispose();
    }

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr param);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr param);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr param);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int max);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
}
