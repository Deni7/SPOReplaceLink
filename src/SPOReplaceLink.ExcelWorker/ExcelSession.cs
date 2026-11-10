using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.ExcelWorker;

/// <summary>
/// One hidden Excel instance in the worker process (XLS-2): no alerts, no link updates, macros disabled,
/// files opened read-only. Excel is late-bound, so any installed version works.
/// </summary>
[SupportedOSPlatform("windows")]
internal sealed class ExcelSession : IDisposable
{
    private const int UpdateLinksNever = 0;
    private const int AutomationSecurityForceDisable = 3;
    private const int LinkTypeExcelLinks = 1;
    private const int CellTypeFormulas = -4123;
    private const int AccessModeNoChange = 1;
    private const int AccessModeExclusive = 3;
    private const int ConflictLocalSessionChanges = 2;
    private const int MaxFormulas = 200_000;

    /// <summary>A password that never matches: a file protected with a password fails to open instead of asking for one.</summary>
    private const string NoPassword = "\u0001x"; // short: xls passwords are at most 15 characters

    /// <summary>Workbook formats Excel may report for a file it opened as a workbook; anything else (text, CSV) is refused.</summary>
    private static readonly HashSet<int> WorkbookFormats = [16, 29, 33, 35, 39, 43, XlFormat.Xls, XlFormat.Xlsb, XlFormat.Xlsx, XlFormat.Xlsm];

    private dynamic? _app;
    private readonly DialogWatch _dialogs;

    private ExcelSession(object app, int pid)
    {
        _app = app;
        Pid = pid;
        _dialogs = new DialogWatch(pid);
    }

    public int Pid { get; }

    /// <summary>Excel showed a dialog during the last request: it is not trusted with another file (it may stay visible).</summary>
    public bool NeedsRestart { get; private set; }

    /// <param name="started">Called with the Excel process id as soon as it is known, before Excel is configured.</param>
    public static ExcelSession Start(Action<int> started)
    {
        var type = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new ExcelException(ExcelErrorKind.Unavailable, L.T("Excel не встановлено (немає Excel.Application)."));
        object app;
        try
        {
            app = Activator.CreateInstance(type)!;
        }
        catch (Exception ex) when (ex is COMException or TargetInvocationException)
        {
            // Often transient (DCOM busy right after a killed Excel): the next file starts a new worker.
            throw new ExcelException(ExcelErrorKind.Crashed, L.T("Excel не запустився: ") + ex.GetBaseException().Message, ex);
        }
        dynamic xl = app;
        GetWindowThreadProcessId(new IntPtr((long)xl.Hwnd), out uint pid);
        var session = new ExcelSession(app, (int)pid);
        started((int)pid);
        try
        {
            xl.Visible = false;
            xl.DisplayAlerts = false;
            xl.AskToUpdateLinks = false;
            xl.AutomationSecurity = AutomationSecurityForceDisable;
            xl.EnableEvents = false;
            xl.ScreenUpdating = false;
            xl.UserControl = false;
        }
        catch
        {
            session.Dispose();
            throw;
        }
        return session;
    }

    /// <summary>
    /// Opens <paramref name="input"/> and saves it as <paramref name="output"/>; reports the format of the input and,
    /// with <paramref name="inspect"/>, the links Excel holds when it saves.
    /// </summary>
    public WorkerResponse Convert(string input, string output, int format, bool inspect) =>
        Watched(() => ConvertWatched(input, output, format, inspect));

    private WorkerResponse ConvertWatched(string input, string output, int format, bool inspect)
    {
        dynamic wb = Open(input);
        try
        {
            int source = (int)wb.FileFormat;
            WorkerResponse seen = inspect ? (WorkerResponse)Collect(wb) : new WorkerResponse(true);
            wb.CheckCompatibility = false;
            string? warning = null;
            try
            {
                SaveAs(wb, output, format, AccessModeNoChange);
            }
            catch (COMException ex) when (IsShared(wb))
            {
                // A legacy shared workbook whose revision log Excel cannot read or apply: Excel itself suggests
                // saving it without sharing. The log is lost, the workbook and its links are kept.
                try
                {
                    SaveAs(wb, output, format, AccessModeExclusive);
                }
                catch (COMException retry)
                {
                    throw new ExcelException(ExcelErrorKind.Excel, L.T("Excel не зберіг файл: ") + ex.Message + L.T(" Без спільного доступу теж: ") + retry.Message, retry);
                }
                warning = Protocol.UnsharedWarning;
            }
            catch (COMException ex)
            {
                throw new ExcelException(ExcelErrorKind.Excel, L.T("Excel не зберіг файл: ") + ex.Message, ex);
            }
            return seen with { SourceFormat = source, Warning = warning };
        }
        finally
        {
            Close(wb);
        }
    }

    private static void SaveAs(dynamic wb, string output, int format, int accessMode)
    {
        if (File.Exists(output))
            File.Delete(output);
        wb.SaveAs(output, format, Missing.Value, Missing.Value, false, false, accessMode, ConflictLocalSessionChanges, false);
    }

    /// <summary>The workbook uses the legacy "Share Workbook" mode.</summary>
    private static bool IsShared(dynamic wb)
    {
        try
        {
            return (bool)wb.MultiUserEditing;
        }
        catch (COMException)
        {
            return false;
        }
    }

    /// <summary>XLS-3: the links Excel sees — Edit Links, hyperlinks of all sheets, formulas with text in them.</summary>
    public WorkerResponse Inspect(string input) => Watched(() =>
    {
        dynamic wb = Open(input);
        try
        {
            return (WorkerResponse)Collect(wb);
        }
        finally
        {
            Close(wb);
        }
    });

    /// <summary>
    /// Test aid: opens the file without a password argument, the way that makes any Excel ask for one,
    /// to check that the question is closed and reported like the prompt some Excel versions show anyway.
    /// </summary>
    public WorkerResponse Prompt(string input) => Watched(() =>
    {
        dynamic wb;
        try
        {
            wb = _app!.Workbooks.Open(input, UpdateLinks: UpdateLinksNever, ReadOnly: true);
        }
        catch (COMException ex)
        {
            throw new ExcelException(ExcelErrorKind.File, L.T("Excel не відкрив файл: ") + ex.Message, ex);
        }
        Close(wb);
        return new WorkerResponse(true);
    });

    /// <summary>Runs a request while <see cref="DialogWatch"/> closes any question Excel asks.</summary>
    private WorkerResponse Watched(Func<WorkerResponse> request)
    {
        NeedsRestart = false;
        _dialogs.Begin();
        try
        {
            return request();
        }
        catch (Exception ex) when (_dialogs.Question is { } question)
        {
            bool password = question.Contains("парол", StringComparison.OrdinalIgnoreCase) // i18n-ignore: matches the text of Excel's dialog
                || question.Contains("password", StringComparison.OrdinalIgnoreCase);
            string what = _dialogs.Killed ? L.T("Excel не закрив вікно і завершений примусово") : L.T("вікно закрито");
            throw new ExcelException(password ? ExcelErrorKind.Password : ExcelErrorKind.File,
                (password ? L.T("Файл захищений паролем: ") : L.T("Excel не обробив файл: ")) + L.F("Excel запитав «{0}», {1}.", question, what), ex);
        }
        finally
        {
            if (_dialogs.End() is not null)
                NeedsRestart = true;
        }
    }

    private static WorkerResponse Collect(dynamic wb)
    {
        var sources = new List<string>();
        var hyperlinks = new List<string>();
        var formulas = new List<string>();
        object? raw = wb.LinkSources(LinkTypeExcelLinks);
        if (raw is Array array)
        {
            foreach (object? s in array)
            {
                if (s is string text)
                    sources.Add(text);
            }
        }

        dynamic sheets = wb.Worksheets;
        int count = sheets.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic ws = sheets.Item(i);
            dynamic links = ws.Hyperlinks;
            int n = links.Count;
            for (int j = 1; j <= n; j++)
            {
                string? address = links.Item(j).Address;
                if (!string.IsNullOrEmpty(address))
                    hyperlinks.Add(address);
            }
            dynamic? cells = null;
            try
            {
                cells = ws.UsedRange.SpecialCells(CellTypeFormulas);
            }
            catch (COMException)
            {
                // No formulas on the sheet.
            }
            if (cells is null)
                continue;
            dynamic areas = cells.Areas;
            int areaCount = areas.Count;
            for (int a = 1; a <= areaCount && formulas.Count < MaxFormulas; a++)
            {
                object? value = areas.Item(a).Formula;
                if (value is string single)
                    AddFormula(formulas, single);
                else if (value is Array grid)
                {
                    foreach (object? f in grid)
                    {
                        if (f is string text)
                            AddFormula(formulas, text);
                    }
                }
            }
        }
        return new WorkerResponse(true, LinkSources: sources, Hyperlinks: hyperlinks, Formulas: formulas);
    }
    private static void AddFormula(List<string> formulas, string formula)
    {
        // Only formulas with string literals can hold a link as text; references to other books are in Edit Links.
        if (formula.StartsWith('=') && formula.Contains('"'))
            formulas.Add(formula);
    }

    private dynamic Open(string input)
    {
        dynamic wb;
        try
        {
            // Named arguments: Missing.Value for skipped positions makes Open fail through the dynamic binder.
            wb = _app!.Workbooks.Open(input, UpdateLinks: UpdateLinksNever, ReadOnly: true, Password: NoPassword,
                WriteResPassword: NoPassword, IgnoreReadOnlyRecommended: true);
        }
        catch (COMException ex)
        {
            bool password = ex.Message.Contains("парол", StringComparison.OrdinalIgnoreCase) // i18n-ignore: matches Excel's error text
                || ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase);
            throw new ExcelException(password ? ExcelErrorKind.Password : ExcelErrorKind.File,
                (password ? L.T("Файл захищений паролем на відкриття: ") : L.T("Excel не відкрив файл: ")) + ex.Message, ex);
        }
        int format = (int)wb.FileFormat;
        if (!WorkbookFormats.Contains(format))
        {
            Close(wb);
            throw new ExcelException(ExcelErrorKind.File, L.F("Excel відкрив файл не як книгу (формат {0}).", format));
        }
        return wb;
    }

    private static void Close(dynamic wb)
    {
        try
        {
            wb.Close(false);
        }
        catch (COMException)
        {
        }
        Marshal.FinalReleaseComObject(wb);
    }

    public void Dispose()
    {
        if (_app is null)
            return;
        _dialogs.Dispose();
        try
        {
            _app.Quit();
        }
        catch (COMException)
        {
        }
        Marshal.FinalReleaseComObject(_app);
        _app = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        // Excel lingers while references are released; never leave it behind (XLS-4).
        try
        {
            using var process = Process.GetProcessById(Pid);
            if (!process.WaitForExit(10_000))
                process.Kill();
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            // Already gone.
        }
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}

/// <summary>Retries calls Excel rejects while it is busy (RPC_E_CALL_REJECTED) instead of failing them.</summary>
[SupportedOSPlatform("windows")]
internal sealed class MessageFilter : IOleMessageFilter
{
    private const int ServerCallRetryLater = 2;

    public static void Register() => _ = CoRegisterMessageFilter(new MessageFilter(), out _);

    int IOleMessageFilter.HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr interfaceInfo) => 0;

    int IOleMessageFilter.RetryRejectedCall(IntPtr callee, int tickCount, int rejectType) =>
        rejectType == ServerCallRetryLater && tickCount < 120_000 ? 200 : -1;

    int IOleMessageFilter.MessagePending(IntPtr callee, int tickCount, int pendingType) => 2;

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);
}

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr interfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr callee, int tickCount, int rejectType);

    [PreserveSig]
    int MessagePending(IntPtr callee, int tickCount, int pendingType);
}
