using System.ComponentModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Components.WebView.Wpf;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.Engine;
using SPOReplaceLink.State;
using SPOReplaceLink.Ui.Components;
using Forms = System.Windows.Forms;

namespace SPOReplaceLink.Desktop;

/// <summary>
/// Main window: the GUI in WebView2 and a tray icon. Closing the window during a run asks whether to keep the run going
/// in the tray or to stop it gracefully (STA-4); a stopped run continues from the same place after the next start.
/// </summary>
internal sealed class MainWindow : Window
{
    private readonly WebApplication _host;
    private readonly RunService _runs;
    private readonly BlazorWebView _webView;
    private readonly Forms.NotifyIcon _tray;
    private readonly string? _selftest;
    private readonly DispatcherTimer _releaseSize;
    private readonly Action _languageChanged;
    private WindowState _shownState = WindowState.Maximized;
    private bool _exiting;
    private bool _trayHintShown;

    // The icon of this assembly, also when another program (the demo tool) hosts the window.
    private static readonly Uri IconUri = new("pack://application:,,,/SPOReplaceLink;component/app.ico");

    public MainWindow(WebApplication host, IServiceProvider services, string dataDir, string? apiWarning, string? selftest)
    {
        _host = host;
        _runs = host.Services.GetRequiredService<RunService>();
        _selftest = selftest;
        Title = "SPOReplaceLink";
        // Opens maximized: a window taller than the screen put the title bar out of sight. The restored size fits the work area too.
        var work = SystemParameters.WorkArea;
        Width = Math.Min(1400, work.Width);
        Height = Math.Min(900, work.Height);
        MinWidth = Math.Min(900, work.Width);
        MinHeight = Math.Min(600, work.Height);
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        WindowState = WindowState.Maximized;
        StateChanged += (_, _) =>
        {
            if (WindowState != WindowState.Minimized)
                _shownState = WindowState;
        };
        // A frame of the icon decoder, not BitmapImage: the window then takes the 16/20/24 px frame for the title bar
        // and a large one for Alt+Tab, instead of scaling the first frame.
        Icon = BitmapFrame.Create(IconUri);

        _webView = new BlazorWebView { HostPage = Path.Combine("wwwroot", "index.html"), Services = services };
        _webView.RootComponents.Add(new RootComponent { Selector = "#app", ComponentType = typeof(Routes) });
        Content = _webView;

        using var iconStream = Application.GetResourceStream(IconUri)!.Stream;
        _tray = new Forms.NotifyIcon
        {
            Icon = new System.Drawing.Icon(iconStream, Forms.SystemInformation.SmallIconSize),
            Text = "SPOReplaceLink",
            Visible = true,
            ContextMenuStrip = new Forms.ContextMenuStrip(),
        };
        BuildTrayMenu(dataDir);
        // The menu follows the language chosen in the app (the event may come from any thread).
        _languageChanged = () => Dispatcher.BeginInvoke(() =>
        {
            if (!_exiting)
                BuildTrayMenu(dataDir);
        });
        L.LanguageChanged += _languageChanged;
        _tray.DoubleClick += (_, _) => Reveal();

        // RDP reconnect changes the display: the WebView2 composition control recreates its D3D device, and a WM_SIZE that
        // arrives while it does (the creation pumps messages) resizes the capture with the released device — an access
        // violation that kills the process (WebView2 SDK 1.0.4258/4314). The control keeps its size until the display settles.
        _releaseSize = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
        _releaseSize.Tick += (_, _) =>
        {
            _releaseSize.Stop();
            _webView.Width = double.NaN;
            _webView.Height = double.NaN;
        };
        SystemEvents.DisplaySettingsChanging += HoldWebViewSize;
        SystemEvents.DisplaySettingsChanged += HoldWebViewSize;

        if (selftest is not null)
            _webView.BlazorWebViewInitialized += async (_, _) => await SelfTestAsync();
        else if (apiWarning is not null)
            Loaded += (_, _) => MessageBox.Show(this, apiWarning, Title, MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private void HoldWebViewSize(object? sender, EventArgs e) => Dispatcher.Invoke(() =>
    {
        if (!_releaseSize.IsEnabled && _webView.ActualWidth > 0 && _webView.ActualHeight > 0)
        {
            _webView.Width = _webView.ActualWidth;
            _webView.Height = _webView.ActualHeight;
        }
        _releaseSize.Stop();
        _releaseSize.Start(); // three seconds after the last change
    });

    public void Reveal()
    {
        Show();
        if (WindowState == WindowState.Minimized)
            WindowState = _shownState; // back to maximized if it was
        Activate();
    }

    private bool RunIsActive(out RunService.ActiveRun? active)
    {
        active = _runs.Active;
        return active is not null && !active.Completion.IsCompleted;
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (_exiting)
        {
            base.OnClosing(e);
            return;
        }
        e.Cancel = true;
        if (!RunIsActive(out var active))
        {
            _ = ExitAsync();
            return;
        }
        var answer = MessageBox.Show(this,
            L.F("Виконується прогін {0}.\n\n", active!.RunId) +
            L.T("Так — згорнути в трей: прогін триватиме, вікно відкривається з іконки біля годинника.\n") +
            L.T("Ні — зупинити прогін і вийти: поточні файли завершаться, решта лишиться в черзі й продовжиться після наступного запуску.\n") +
            L.T("Скасувати — лишитися у вікні."),
            Title, MessageBoxButton.YesNoCancel, MessageBoxImage.Question, MessageBoxResult.Yes);
        if (answer == MessageBoxResult.Yes)
            HideToTray();
        else if (answer == MessageBoxResult.No)
            _ = ExitAsync();
    }

    private void BuildTrayMenu(string dataDir)
    {
        var menu = _tray.ContextMenuStrip!;
        menu.Items.Clear();
        menu.Items.Add(L.T("Відкрити"), null, (_, _) => Reveal());
        menu.Items.Add(L.F("Папка даних: {0}", dataDir), null, (_, _) => System.Diagnostics.Process.Start("explorer.exe", dataDir));
        menu.Items.Add(new Forms.ToolStripSeparator());
        menu.Items.Add(L.T("Вийти"), null, (_, _) => RequestExit());
    }

    private void HideToTray()
    {
        Hide();
        if (!_trayHintShown)
        {
            _trayHintShown = true;
            _tray.ShowBalloonTip(5000, "SPOReplaceLink", L.T("Прогін триває у фоні. Відкрити вікно — подвійний клік по іконці."), Forms.ToolTipIcon.Info);
        }
    }

    private void RequestExit()
    {
        if (RunIsActive(out var active)
            && MessageBox.Show(this, L.F("Зупинити прогін {0} і вийти? Поточні файли завершаться, решта продовжиться після наступного запуску.", active!.RunId),
                Title, MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            return;
        _ = ExitAsync();
    }

    /// <summary>Graceful stop of the active run (up to 30 s for current files), then the engine and the window.</summary>
    private async Task ExitAsync(int code = 0)
    {
        if (_exiting)
            return;
        _exiting = true;
        L.LanguageChanged -= _languageChanged;
        Title = L.T("SPOReplaceLink — зупинка…");
        IsEnabled = false;
        try
        {
            await _host.StopAsync();
            await _host.DisposeAsync();
        }
        finally
        {
            SystemEvents.DisplaySettingsChanging -= HoldWebViewSize;
            SystemEvents.DisplaySettingsChanged -= HoldWebViewSize;
            _releaseSize.Stop();
            _tray.Visible = false;
            _tray.Dispose();
            Close();
            Application.Current.Shutdown(code);
        }
    }

    /// <summary>
    /// --selftest: the window starts, the pages render and the main sections open; the result goes to a file.
    /// Lets the build be checked on a machine without anyone looking at the screen.
    /// </summary>
    private async Task SelfTestAsync()
    {
        var log = new List<string>();
        bool ok = true;
        async Task<string> Eval(string js)
        {
            string raw = await _webView.WebView.CoreWebView2.ExecuteScriptAsync(js);
            return JsonSerializer.Deserialize<string>(raw) ?? string.Empty;
        }
        async Task<bool> WaitFor(string what, string js)
        {
            for (int i = 0; i < 120; i++)
            {
                if (await Eval(js) is { Length: > 0 } found)
                {
                    log.Add($"ok   {what}: {found}");
                    return true;
                }
                await Task.Delay(250);
            }
            log.Add($"FAIL {what}");
            return false;
        }

        ok &= await WaitFor("дашборд", "document.querySelector('.mud-appbar') ? document.querySelector('.mud-main-content').innerText.substring(0, 60).replace(/\\s+/g, ' ') : ''"); // i18n-ignore: self-test, runs in Ukrainian
        ok &= await WaitFor("стилі MudBlazor", "getComputedStyle(document.querySelector('.mud-appbar')).position === 'fixed' ? 'застосовано' : ''"); // i18n-ignore: self-test, runs in Ukrainian
        // The home page shows the last run, when the data folder has one: its charts (uPlot, through JS) draw and a picture is saved.
        if (await Eval("document.querySelector('.mud-main-content').innerText.trimStart().startsWith('Прогін') ? 'yes' : ''") == "yes") // i18n-ignore: self-test, runs in Ukrainian
        {
            ok &= await WaitFor("графіки прогону", "document.querySelectorAll('.uplot').length >= 2 ? 'намальовано' : (document.querySelector('.mud-main-content').innerText.includes('Графіки з\\'являться') ? 'ще немає знімків' : '')"); // i18n-ignore: self-test, runs in Ukrainian
            await CaptureAsync(Path.ChangeExtension(_selftest!, ".dashboard.png"));
            ok &= await WaitFor("розбивка за цілями", "(() => { const r = document.querySelector('.breakdown-row'); if (r) r.scrollIntoView(); return r ? 'є' : ''; })()"); // i18n-ignore: self-test, runs in Ukrainian
            await CaptureAsync(Path.ChangeExtension(_selftest!, ".targets.png"));
        }
        foreach (var (href, heading) in new[] { ("runs", "Прогони"), ("profiles", "Профілі прогону"), ("templates", "Шаблони правил"), ("connections", "Підключення до SharePoint") }) // i18n-ignore: self-test, runs in Ukrainian
        {
            await Eval($"(document.querySelector('a[href=\"{href}\"]') || {{ click(){{}} }}).click(), ''");
            ok &= await WaitFor(href, $"document.querySelector('.mud-main-content').innerText.includes('{heading}') ? '{heading}' : ''");
        }

        // Help: the guide opens by sections; a picture of a section is saved next to the result for a look later.
        await Eval("(document.querySelector('a[href=\"help\"]') || { click(){} }).click(), ''");
        ok &= await WaitFor("довідка", "document.querySelector('.help-content') && document.querySelector('.help-nav') ? 'Про інструкцію' : ''"); // i18n-ignore: self-test, runs in Ukrainian
        await Eval("(() => { const a = document.createElement('a'); a.href = 'help/' + encodeURIComponent('вкладка-продуктивність'); document.body.appendChild(a); a.click(); a.remove(); return ''; })()"); // i18n-ignore: self-test, runs in Ukrainian
        ok &= await WaitFor("довідка: розділ", "document.querySelector('.mud-main-content').innerText.includes('performance.budgetPercent') ? 'Вкладка «Продуктивність»' : ''"); // i18n-ignore: self-test, runs in Ukrainian
        await CaptureAsync(Path.ChangeExtension(_selftest!, ".help.png"));
        await Eval("(document.querySelector('a[href=\"profiles\"]') || { click(){} }).click(), ''");
        await Eval("(() => { const a = document.createElement('a'); a.href = 'profiles/~new'; document.body.appendChild(a); a.click(); a.remove(); return ''; })()");
        ok &= await WaitFor("кнопка довідки на вкладці", "[...document.querySelectorAll('button')].some(b => b.textContent.includes('Що означають ці поля')) ? 'є' : ''"); // i18n-ignore: self-test, runs in Ukrainian
        await Eval("([...document.querySelectorAll('button')].find(b => b.textContent.includes('Що означають ці поля')) || { click(){} }).click(), ''"); // i18n-ignore: self-test, runs in Ukrainian
        ok &= await WaitFor("вікно довідки", "document.querySelector('.mud-dialog .help-content') ? 'відкрито' : ''"); // i18n-ignore: self-test, runs in Ukrainian
        await CaptureAsync(Path.ChangeExtension(_selftest!, ".help-dialog.png"));
        log.Add(ok ? "RESULT ok" : "RESULT fail");

        async Task CaptureAsync(string path)
        {
            await Task.Delay(500); // let the page settle
            await using var png = File.Create(path);
            await _webView.WebView.CoreWebView2.CapturePreviewAsync(Microsoft.Web.WebView2.Core.CoreWebView2CapturePreviewImageFormat.Png, png);
            log.Add($"ok   знімок: {path}"); // i18n-ignore: self-test, runs in Ukrainian
        }
        await File.WriteAllLinesAsync(_selftest!, log);
        await ExitAsync(ok ? 0 : 1);
    }
}
