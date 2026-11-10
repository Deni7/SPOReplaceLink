using SPOReplaceLink.Engine;
using SPOReplaceLink.OpenXml;
using SPOReplaceLink.State;

namespace SPOReplaceLink.Ui.Services;

/// <summary>
/// Where the GUI puts a file the user exports (GUI-10): the desktop app asks with "Save as";
/// tests and other hosts may write elsewhere.
/// </summary>
public interface IFileSaver
{
    /// <param name="suggestedName">File name offered in the dialog, e.g. run-12-files.xlsx.</param>
    /// <returns>Where the file was saved, or null if the user cancelled.</returns>
    Task<string?> SaveAsync(string suggestedName, Func<Stream, Task> write);
}

/// <summary>Exports shared by the GUI and the REST API.</summary>
public static class Exports
{
    public const string XlsxType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /// <summary>GUI-10: the files table with a filter, as CSV or XLSX. Read in id order with keyset paging.</summary>
    public static void WriteFiles(StateStore store, long runId, FileFilter filter, string format, Stream output)
    {
        var rows = ReportExporter.FileRows(AllFiles(store, runId, filter));
        if (format == "xlsx")
            SimpleXlsxWriter.Write(output, "files", rows);
        else
            ReportExporter.WriteCsv(output, rows);
    }

    /// <summary>A run report (files, links, residuals, attention, prefixes, summary) as CSV/TXT or XLSX.</summary>
    public static void WriteReport(StateStore store, long runId, string name, string format, Stream output, int prefixDepth = ReportExporter.DefaultPrefixDepth)
    {
        if (format == "xlsx")
            ReportExporter.WriteXlsx(store, runId, name, output, prefixDepth);
        else
            ReportExporter.Write(store, runId, name, output, prefixDepth);
    }

    public static string ReportFileName(long runId, string name, string format) =>
        $"run-{runId}-{name}.{(format == "xlsx" ? "xlsx" : ReportExporter.IsText(name) ? "txt" : "csv")}";

    public static IEnumerable<FileRow> AllFiles(StateStore store, long runId, FileFilter filter)
    {
        long after = 0;
        while (true)
        {
            var page = store.QueryFilesAfter(runId, filter, after, 5000);
            foreach (var f in page)
                yield return f;
            if (page.Count < 5000)
                yield break;
            after = page[^1].Id;
        }
    }
}
