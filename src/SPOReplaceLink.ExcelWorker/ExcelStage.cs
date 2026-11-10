using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Urls;
using SPOReplaceLink.OpenXml;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.ExcelWorker;

/// <summary>
/// XLS-3, XLS-5: xls, and xlsb with old prefixes left in binary records, go through Excel.
/// Excel saves a copy as xlsm, the links in it are replaced the same way as in any xlsx (rules, places, self-check),
/// and Excel saves the result back in the original format. Links are never changed with Excel's own ChangeLink:
/// it reads the new source at once, and a source Excel cannot reach (SharePoint) turns the cached values into #REF!.
/// Then the written file is opened in Excel once more and every link Excel reports is checked against the rules.
/// </summary>
public sealed class ExcelStage(WorkbookProcessor processor, IExcelConverter excel)
{
    /// <summary>Excel cannot open longer paths.</summary>
    private const int MaxExcelPath = 218;

    public static bool IsNeeded(FileResult result) => result.Status == FileStatus.NeedsExcel || result.NeedsExcel;

    /// <param name="workDir">An empty folder for this file; deleted afterwards.</param>
    public async Task<FileResult> ProcessAsync(string inputPath, string outputPath, string? documentFolder, string displayPath, string workDir, CancellationToken ct)
    {
        string ext = Path.GetExtension(inputPath).ToLowerInvariant();
        var result = new FileResult { Path = displayPath, Format = ext.TrimStart('.') };
        try
        {
            // A folder left by an earlier attempt that could not be cleaned up must not fail this one.
            TryDeleteDirectory(workDir);
            var local = LocalCopy.Create(workDir);
            string book = Path.Combine(local.Folder, "book" + ext);
            string mid = Path.Combine(local.Folder, "book.xlsm");
            string midOut = Path.Combine(local.Folder, "book-out.xlsm");
            string final = Path.Combine(local.Folder, "book-out" + ext);
            if (midOut.Length > MaxExcelPath)
                return Fail(result, ErrorClass.Internal, L.F("Тимчасова папка задовга для Excel ({0} символів, Excel приймає до {1}).", midOut.Length, MaxExcelPath));
            File.Copy(inputPath, book);

            var converted = await excel.ConvertAsync(book, mid, XlFormat.Xlsm, ct).ConfigureAwait(false);
            int source = converted.SourceFormat ?? 0;
            // Sharing is turned off already in the copy; the file is uploaded only if links changed.
            string? warning = converted.Warning;
            var changed = processor.Process(mid, midOut, documentFolder, displayPath, local);
            result.Links.AddRange(changed.Links);
            if (changed.Status == FileStatus.Failed)
                return Fail(result, changed.ErrorClass ?? ErrorClass.Internal, changed.Error ?? L.T("Помилка обробки копії xlsm."));

            string written = book;
            if (changed.OutputWritten)
            {
                var saved = await excel.ConvertInspectingAsync(midOut, final, source > 0 ? source : XlFormat.ForExtension(ext) ?? XlFormat.Xls, ct).ConfigureAwait(false);
                if (LocalOnSameDrive(saved, local) is { } risky)
                {
                    // Excel saves links to files on the drive of the folder it saves in without the drive letter;
                    // opened from SharePoint such a link would point elsewhere. Nothing is written.
                    return Fail(result, ErrorClass.Validation,
                        L.F("Посилання на локальний диск {0} ({1}): Excel збереже його без літери диска. Обробіть файл з тимчасовою папкою на іншому диску (папка даних застосунку або --temp у local).",
                        Path.GetPathRoot(local.Folder), risky));
                }
                written = final;
                warning ??= saved.Warning;
            }

            // What Excel itself sees in the file that would be uploaded.
            var seen = await excel.InspectAsync(written, ct).ConfigureAwait(false);
            var residuals = new List<Residual>(processor.CheckLinks(Seen(seen), documentFolder, local));
            if (ext == ".xls")
            {
                residuals.AddRange(processor.ScanFile(written, "xls"));
            }
            else
            {
                string scratch = Path.Combine(local.Folder, "check" + ext);
                var check = processor.Process(written, scratch, documentFolder, displayPath, local);
                residuals.AddRange(check.Residuals);
            }
            result.Residuals.AddRange(residuals.DistinctBy(r => (r.Part, r.Match, r.Context)));

            if (changed.OutputWritten)
            {
                File.Copy(final, outputPath, overwrite: true);
                result.ChangedParts.Add("excel");
                result.OutputWritten = true;
                result.Status = FileStatus.Planned;
                result.Error = warning;
            }
            else
            {
                result.Status = FileStatus.Unchanged;
            }
            return result;
        }
        catch (ExcelException ex) when (ex.Kind == ExcelErrorKind.Password)
        {
            result.Status = FileStatus.SkippedEncrypted;
            result.Error = ex.Message;
            return result;
        }
        catch (ExcelException ex) when (ex.Kind == ExcelErrorKind.File)
        {
            return Fail(result, ErrorClass.File, ex.Message);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Fail(result, ErrorClass.File, ex.Message);
        }
        finally
        {
            TryDeleteDirectory(workDir);
        }
    }

    /// <summary>
    /// A link Excel holds when saving that is an absolute path on the drive of the temporary folder, or null.
    /// Locations inside the copy's root were relative to the document and stay relative.
    /// </summary>
    private static string? LocalOnSameDrive(WorkerResponse saved, LocalCopy local)
    {
        string drive = Path.GetPathRoot(local.Folder)!.TrimEnd('\\');
        foreach (string value in (saved.LinkSources ?? []).Concat(saved.Hyperlinks ?? []))
        {
            string path = value.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ? value[5..].TrimStart('/', '\\') : value;
            if (LinkClassifier.Classify(path) != LinkForm.Drive || !path.StartsWith(drive, StringComparison.OrdinalIgnoreCase))
                continue;
            string full = Path.GetFullPath(path);
            if (!full.StartsWith(local.Root.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase))
                return value;
        }
        return null;
    }
    private static IEnumerable<(string, LinkKind, string)> Seen(WorkerResponse seen)
    {
        foreach (string s in seen.LinkSources ?? [])
            yield return ("excel:link_sources", LinkKind.LinkSource, s);
        foreach (string s in seen.Hyperlinks ?? [])
            yield return ("excel:hyperlinks", LinkKind.Hyperlink, s);
        foreach (string s in seen.Formulas ?? [])
            yield return ("excel:formulas", LinkKind.Formula, s);
    }

    private static FileResult Fail(FileResult result, ErrorClass errorClass, string message)
    {
        result.Status = FileStatus.Failed;
        result.ErrorClass = errorClass;
        result.Error = message;
        return result;
    }

    private static void TryDeleteDirectory(string path)
    {
        for (int i = 0; i < 3; i++)
        {
            try
            {
                if (Directory.Exists(path))
                    Directory.Delete(path, true);
                return;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Excel may hold the file for a moment after closing it.
                Thread.Sleep(200);
            }
        }
    }
}
