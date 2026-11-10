using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.ExcelWorker;

/// <summary>
/// The worker process (XLS-1): drives one Excel instance and works only with local copies of files.
/// Requests come as JSON lines on stdin, answers go to stdout (see <see cref="Protocol"/>).
/// <c>--fake</c> works without Excel, for tests of the pool and the engine: a "fake xls" is the OLE signature followed by
/// xlsx bytes, convert adds or strips the signature, inspect reports nothing; names with hang / crash / secret / prompt misbehave.
/// </summary>
internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        var input = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
        var output = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
        bool fake = args.Contains("--fake");
        L.SetLanguage(L.Normalize(Environment.GetEnvironmentVariable("SPO_RL_LANG")) ?? L.FromSystem());
        if (!fake && !OperatingSystem.IsWindows())
        {
            Write(output, new WorkerReady(false, 0, L.T("Excel працює лише у Windows."), ExcelErrorKind.Unavailable));
            return 2;
        }
        return fake ? Run(input, output, null) : RunWithExcel(input, output);
    }

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    private static int RunWithExcel(StreamReader input, StreamWriter output)
    {
        MessageFilter.Register();
        ExcelSession session;
        try
        {
            session = ExcelSession.Start(pid => Write(output, new WorkerReady(false, pid)));
        }
        catch (ExcelException ex)
        {
            Write(output, new WorkerReady(false, 0, ex.Message, ex.Kind));
            return 1;
        }
        catch (Exception ex)
        {
            Write(output, new WorkerReady(false, 0, L.T("Excel не запустився: ") + ex.Message, ExcelErrorKind.Crashed));
            return 1;
        }
        using (session)
            return Run(input, output, session);
    }

    private static int Run(StreamReader input, StreamWriter output, object? session)
    {
        Write(output, new WorkerReady(true, session is null ? 0 : Pid(session)));
        while (input.ReadLine() is { } line)
        {
            WorkerRequest? request;
            try
            {
                request = JsonSerializer.Deserialize<WorkerRequest>(line, Protocol.Json);
            }
            catch (JsonException ex)
            {
                Write(output, new WorkerResponse(false, L.T("Некоректний запит: ") + ex.Message, ExcelErrorKind.Excel));
                continue;
            }
            switch (request?.Op)
            {
                case "quit" or null:
                    return 0;
                case "stall":
                    // Test aid: behaves like a hung Excel — never answers.
                    Thread.Sleep(Timeout.Infinite);
                    break;
                case "exit":
                    // Test aid: behaves like a crashed worker.
                    Environment.Exit(3);
                    break;
            }
            Write(output, Handle(request!, session));
        }
        return 0;
    }

    private static WorkerResponse Handle(WorkerRequest request, object? session)
    {
        var response = Answer(request, session);
        return session is not null && OperatingSystem.IsWindows() && ((ExcelSession)session).NeedsRestart
            ? response with { Restart = true }
            : response;
    }

    private static WorkerResponse Answer(WorkerRequest request, object? session)
    {
        try
        {
            if (session is null)
                return Fake(request);
            if (!OperatingSystem.IsWindows())
                return new WorkerResponse(false, L.T("Excel працює лише у Windows."), ExcelErrorKind.Unavailable);
            return Real(request, session);
        }
        catch (ExcelException ex)
        {
            return new WorkerResponse(false, ex.Message, ex.Kind);
        }
        catch (COMException ex) when ((uint)ex.HResult is 0x800706BA or 0x800706BE or 0x80010108)
        {
            // RPC server unavailable / call failed / object disconnected: Excel itself is gone.
            return new WorkerResponse(false, L.T("Excel завершився: ") + ex.Message, ExcelErrorKind.Crashed);
        }
        catch (Exception ex)
        {
            return new WorkerResponse(false, ex.GetType().Name + ": " + ex.Message, ExcelErrorKind.Excel);
        }
    }

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    private static WorkerResponse Real(WorkerRequest request, object session)
    {
        var excel = (ExcelSession)session;
        return request.Op switch
        {
            "convert" => excel.Convert(request.Input!, request.Output!, request.Format, request.Inspect),
            "inspect" => excel.Inspect(request.Input!),
            "prompt" => excel.Prompt(request.Input!),
            _ => new WorkerResponse(false, L.F("Невідома операція «{0}».", request.Op), ExcelErrorKind.Excel),
        };
    }

    private static WorkerResponse Fake(WorkerRequest request)
    {
        if (Marked(request, "hang"))
            Thread.Sleep(Timeout.Infinite);
        if (Marked(request, "crash"))
            Environment.Exit(3);
        if (Marked(request, "secret"))
            return new WorkerResponse(false, L.T("Файл захищений паролем на відкриття."), ExcelErrorKind.Password);
        if (Marked(request, "prompt"))
            return new WorkerResponse(false, L.T("Файл захищений паролем: Excel запитав «Пароль», вікно закрито."), ExcelErrorKind.Password, Restart: true);
        return request.Op switch
        {
            "convert" => Copy(request),
            "inspect" => new WorkerResponse(true, LinkSources: [], Hyperlinks: [], Formulas: []),
            _ => new WorkerResponse(false, L.F("Невідома операція «{0}».", request.Op), ExcelErrorKind.Excel),
        };
    }

    /// <summary>The marker in the file name, or as "fake-marker" in the bytes (e.g. the name of a zip entry: the engine renames copies).</summary>
    private static bool Marked(WorkerRequest request, string marker)
    {
        string path = request.Input ?? string.Empty;
        if (Path.GetFileName(path).Contains(marker, StringComparison.OrdinalIgnoreCase))
            return true;
        return File.Exists(path) && File.ReadAllBytes(path).AsSpan().IndexOf(Encoding.ASCII.GetBytes("fake-" + marker)) >= 0;
    }

    private static readonly byte[] OleSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    private static WorkerResponse Copy(WorkerRequest request)
    {
        byte[] bytes = File.ReadAllBytes(request.Input!);
        bool ole = bytes.AsSpan().StartsWith(OleSignature);
        if (ole && request.Format != XlFormat.Xls)
            bytes = bytes[OleSignature.Length..];
        else if (!ole && request.Format == XlFormat.Xls)
            bytes = [.. OleSignature, .. bytes];
        File.WriteAllBytes(request.Output!, bytes);
        return new WorkerResponse(true, SourceFormat: XlFormat.ForExtension(Path.GetExtension(request.Input!)) ?? 0,
            LinkSources: request.Inspect ? [] : null, Hyperlinks: request.Inspect ? [] : null, Formulas: request.Inspect ? [] : null,
            Warning: Marked(request, "shared") ? Protocol.UnsharedWarning : null);
    }

    private static int Pid(object session) => OperatingSystem.IsWindows() ? ((ExcelSession)session).Pid : 0;

    private static void Write<T>(StreamWriter output, T message) => output.WriteLine(JsonSerializer.Serialize(message, Protocol.Json));
}
