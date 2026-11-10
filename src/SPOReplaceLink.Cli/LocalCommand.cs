using System.Diagnostics;
using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Reports;
using SPOReplaceLink.Core.Urls;
using SPOReplaceLink.ExcelWorker;
using SPOReplaceLink.OpenXml;

namespace SPOReplaceLink.Cli;

internal sealed record LocalOptions(
    string Input,
    string Config,
    string? Out,
    bool Write,
    string? BaseUrl,
    string? Report,
    int Parallel,
    string Delimiter,
    int? Excel = null,
    string? Temp = null);

/// <summary>
/// <c>local</c>: processes a folder on disk. The input is never modified; changed copies go to --out
/// only with --write (P-7), unchanged files are not copied.
/// </summary>
internal static class LocalCommand
{
    public static async Task<int> RunAsync(LocalOptions o, CancellationToken ct)
    {
        var loaded = ProfileLoader.Load(o.Config);
        if (loaded is null)
            return ExitCodes.InvalidInput;
        var (profile, rules) = loaded.Value;

        string? error = CheckOptions(o, out string inputRoot, out bool singleFile);
        if (error is not null)
        {
            Console.Error.WriteLine(L.T("ПОМИЛКА: ") + error);
            return ExitCodes.InvalidInput;
        }

        var files = singleFile
            ? [Path.GetFullPath(o.Input)]
            : Directory.EnumerateFiles(inputRoot, "*", SearchOption.AllDirectories)
                .Where(profile.Filters.Accepts)
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

        string reportDir = o.Report ?? Path.Combine(Environment.CurrentDirectory, "reports", DateTime.Now.ToString("yyyyMMdd-HHmmss"));
        string tempRoot = o.Temp ?? Path.Combine(Path.GetTempPath(), "spo-replace-link");
        string tempDir = Path.Combine(tempRoot, Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(tempDir);

        Console.WriteLine(L.F("Профіль: {0}; правил: {1}; файлів: {2}; ", profile.Name ?? Path.GetFileName(o.Config), rules.Rules.Count, files.Count) +
            (o.Write ? L.F("запис у {0}", o.Out) : L.T("dry-run (без запису)")));

        var processor = new WorkbookProcessor(rules, profile.Places);
        // XLS-4: xls and xlsb residuals through Excel; --excel overrides the profile, 0 switches Excel off.
        int excelProcesses = o.Excel ?? profile.Performance.ExcelProcesses;
        await using var excel = excelProcesses > 0 && OperatingSystem.IsWindows()
            ? new ExcelPool(new ExcelPoolOptions(excelProcesses, TimeSpan.FromMinutes(profile.Performance.ExcelTimeoutMinutes), tempRoot))
            : null;
        var stage = excel is null ? null : new ExcelStage(processor, excel);
        long maxBytes = profile.Filters.MaxFileSizeMb * 1024L * 1024L;
        var results = new List<FileResult>();
        var sync = new object();
        var sw = Stopwatch.StartNew();

        try
        {
            using var report = new CsvReport(reportDir, o.Delimiter[0]);
            await Parallel.ForEachAsync(files, new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, o.Parallel), CancellationToken = ct },
                async (file, token) =>
                {
                    string relative = singleFile ? Path.GetFileName(file) : Path.GetRelativePath(inputRoot, file);
                    FileResult result;
                    try
                    {
                        result = await ProcessOneAsync(processor, stage, file, relative, o, tempDir, maxBytes, token);
                    }
                    catch (ExcelException ex) when (ex.Kind == ExcelErrorKind.Unavailable)
                    {
                        result = new FileResult
                        {
                            Path = relative,
                            Format = Path.GetExtension(file).TrimStart('.').ToLowerInvariant(),
                            Status = FileStatus.NeedsExcel,
                            Error = L.F("Excel недоступний: {0}", ex.Message),
                        };
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ExcelException)
                    {
                        // One file that cannot be read or written must not stop the run (P-5).
                        result = new FileResult
                        {
                            Path = relative,
                            Format = Path.GetExtension(file).TrimStart('.').ToLowerInvariant(),
                            Status = FileStatus.Failed,
                            ErrorClass = ex is ExcelException ? ErrorClass.Excel : ErrorClass.File,
                            Error = ex.Message,
                        };
                    }
                    report.Add(result);
                    lock (sync)
                    {
                        results.Add(result);
                        Console.WriteLine($"{result.StatusCode,-22} {relative}" +
                            (result.Replacements > 0 ? L.F("  замін: {0}", result.Replacements) : string.Empty) +
                            (result.Residuals.Count > 0 ? L.F("  залишків: {0}", result.Residuals.Count) : string.Empty) +
                            (result.Error is not null ? $"  — {result.Error}" : string.Empty));
                    }
                });
        }
        finally
        {
            try
            {
                Directory.Delete(tempDir, recursive: true);
            }
            catch (IOException)
            {
            }
        }

        PrintSummary(results, sw.Elapsed, reportDir);
        return results.Any(r => r.NeedsAttention) ? ExitCodes.Attention : ExitCodes.Ok;
    }

    private static async Task<FileResult> ProcessOneAsync(WorkbookProcessor processor, ExcelStage? excel, string file, string relative, LocalOptions o,
        string tempDir, long maxBytes, CancellationToken ct)
    {
        string ext = Path.GetExtension(file);
        if (maxBytes > 0 && new FileInfo(file).Length > maxBytes)
        {
            return new FileResult
            {
                Path = relative,
                Format = ext.TrimStart('.').ToLowerInvariant(),
                Status = FileStatus.SkippedTooLarge,
                Error = L.T("Більший за filters.maxFileSizeMb."),
            };
        }

        string? folder = o.BaseUrl is null ? null : CombineFolder(o.BaseUrl, Path.GetDirectoryName(relative));
        string temp = Path.Combine(tempDir, Guid.NewGuid().ToString("N") + ext);
        var result = processor.Process(file, temp, folder, relative);
        if (excel is not null && ExcelStage.IsNeeded(result))
            result = await excel.ProcessAsync(file, temp, folder, relative, Path.Combine(tempDir, Guid.NewGuid().ToString("N")[..8]), ct);
        if (!result.OutputWritten)
            return result;

        if (o.Write)
        {
            string destination = Path.Combine(o.Out!, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            File.Move(temp, destination, overwrite: true);
            result.Status = FileStatus.Done;
        }
        else
        {
            File.Delete(temp);
        }
        return result;
    }

    /// <summary>Old location of a subfolder: base URL + the local relative folder.</summary>
    internal static string CombineFolder(string baseUrl, string? relativeDir)
    {
        if (string.IsNullOrEmpty(relativeDir))
            return baseUrl;
        bool http = LinkClassifier.Classify(baseUrl) == LinkForm.Http;
        char separator = http ? '/' : '\\';
        var segments = relativeDir.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            .Where(s => s.Length > 0)
            .Select(s => http ? UrlText.EncodeUnsafe(s) : s);
        return baseUrl.TrimEnd('/', '\\') + separator + string.Join(separator, segments);
    }

    private static string? CheckOptions(LocalOptions o, out string inputRoot, out bool singleFile)
    {
        inputRoot = Path.GetFullPath(o.Input);
        singleFile = File.Exists(inputRoot);
        if (!singleFile && !Directory.Exists(inputRoot))
            return L.F("Не знайдено папку або файл «{0}».", o.Input);
        if (singleFile)
            inputRoot = Path.GetDirectoryName(inputRoot)!;
        if (o.Delimiter.Length != 1)
            return L.T("--csv-delimiter має бути одним символом.");
        if (o.Write && o.Out is null)
            return L.T("--write потребує --out: вихідні файли не змінюються на місці.");
        if (o.Out is not null)
        {
            string outFull = Path.GetFullPath(o.Out).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            string inFull = inputRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (outFull.StartsWith(inFull, StringComparison.OrdinalIgnoreCase) || inFull.StartsWith(outFull, StringComparison.OrdinalIgnoreCase))
                return L.T("--out не може бути всередині вхідної папки або містити її.");
        }
        if (o.BaseUrl is not null && !LinkClassifier.IsAbsolute(LinkClassifier.Classify(o.BaseUrl)))
            return L.T("--base-url має бути повною адресою (https://…, \\\\сервер\\…, C:\\…).");
        return null;
    }

    private static void PrintSummary(List<FileResult> results, TimeSpan elapsed, string reportDir)
    {
        Console.WriteLine();
        Console.WriteLine(L.F("Оброблено файлів: {0} за {1:hh\\:mm\\:ss}", results.Count, elapsed));
        foreach (var group in results.GroupBy(r => r.StatusCode).OrderBy(g => g.Key, StringComparer.Ordinal))
            Console.WriteLine($"  {group.Key,-22} {group.Count()}");
        Console.WriteLine(L.F("Замін посилань: {0}", results.Sum(r => r.Replacements)));
        int attention = results.Count(r => r.NeedsAttention);
        if (attention > 0)
            Console.WriteLine(L.F("Потребують уваги: {0} (див. files.csv і residuals.csv)", attention));
        Console.WriteLine(L.F("Звіти: {0}", reportDir));
    }
}
