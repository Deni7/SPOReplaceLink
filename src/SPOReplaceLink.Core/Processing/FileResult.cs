namespace SPOReplaceLink.Core.Processing;

/// <summary>Where in the workbook a link was found.</summary>
public enum LinkKind
{
    Hyperlink,
    LinkSource,
    Formula,
    /// <summary>Cached text result of a formula (&lt;v&gt;) that is itself a link.</summary>
    CachedValue,
    OtherExternal,
    /// <summary>
    /// Folder the workbook was last saved in (x15ac:absPath). Metadata, not a link: updated only when the file
    /// is written anyway, not counted as a replacement and not reported as a residual.
    /// </summary>
    WorkbookPath,
}

public enum LinkAction
{
    Replaced,
    /// <summary>No rule matched — feeds the prefix histogram of unknown locations.</summary>
    NoRule,
    /// <summary>Already points to the new location.</summary>
    Current,
    /// <summary>Relative link without the document's old location.</summary>
    Unresolved,
    /// <summary>Replacement in this place is switched off in the profile.</summary>
    PlaceDisabled,
}

/// <summary>One link found in a part of the workbook (REP-2).</summary>
public sealed record LinkRecord(
    string Part,
    LinkKind Kind,
    string OldValue,
    string? NewValue,
    LinkAction Action,
    int? RuleIndex,
    string? ResolvedFrom);

/// <summary>An old prefix still present after replacement (RES-1, RES-2).</summary>
public sealed record Residual(string Part, string Encoding, string Match, string Context);

public enum FileStatus
{
    /// <summary>Nothing to replace.</summary>
    Unchanged,
    /// <summary>Changes found and verified; not written (dry-run).</summary>
    Planned,
    /// <summary>Changes written.</summary>
    Done,
    /// <summary>Format needs Excel (xls); handled by ExcelWorker in stage 4.</summary>
    NeedsExcel,
    SkippedEncrypted,
    SkippedTooLarge,
    Failed,
}

/// <summary>Error classes from FEATURES 9.1.</summary>
public enum ErrorClass
{
    /// <summary>Network, 5xx, 429 after all request retries, timeout. Retried in the run.</summary>
    Transient,
    /// <summary>412 / changed during processing. Retried with a fresh download.</summary>
    Conflict,
    /// <summary>423, checked out, co-authoring lock. Deferred.</summary>
    Locked,
    /// <summary>Excel timeout or crash (ExcelWorker).</summary>
    Excel,
    /// <summary>401 after token refresh, 403.</summary>
    Auth,
    /// <summary>Damaged or unrecognized file.</summary>
    File,
    /// <summary>Self-check of the result failed (OOX-14); nothing was written.</summary>
    Validation,
    /// <summary>Bug in the tool.</summary>
    Internal,
}

public sealed class FileResult
{
    public required string Path { get; init; }

    public required string Format { get; init; }

    public FileStatus Status { get; set; }

    /// <summary>xlsb: old prefixes left in binary parts, Excel is needed to finish (XLB-2).</summary>
    public bool NeedsExcel { get; set; }

    /// <summary>A verified modified copy was written to the output path.</summary>
    public bool OutputWritten { get; set; }

    public ErrorClass? ErrorClass { get; set; }

    public string? Error { get; set; }

    public List<LinkRecord> Links { get; } = [];

    public List<Residual> Residuals { get; } = [];

    public List<string> ChangedParts { get; } = [];

    public int Replacements => Links.Count(l => l.Action == LinkAction.Replaced && l.Kind != LinkKind.WorkbookPath);

    public int ReplacementsOf(LinkKind kind) => Links.Count(l => l.Action == LinkAction.Replaced && l.Kind == kind);

    /// <summary>Status as shown in reports: unchanged, planned, done, done:residual, done:needs_excel, skipped:encrypted, failed:file…</summary>
    public string StatusCode
    {
        get
        {
            string code = Status switch
            {
                FileStatus.Unchanged => "unchanged",
                FileStatus.Planned => "planned",
                FileStatus.Done => "done",
                FileStatus.NeedsExcel => "needs_excel",
                FileStatus.SkippedEncrypted => "skipped:encrypted",
                FileStatus.SkippedTooLarge => "skipped:too_large",
                FileStatus.Failed => "failed:" + (ErrorClass ?? Processing.ErrorClass.Internal).ToString().ToLowerInvariant(),
                _ => Status.ToString(),
            };
            if (Status is FileStatus.Unchanged or FileStatus.Planned or FileStatus.Done)
            {
                if (NeedsExcel)
                    code += ":needs_excel";
                else if (Residuals.Count > 0)
                    code += ":residual";
            }
            return code;
        }
    }

    /// <summary>The file needs someone to look at it: failed, residuals left, or waits for Excel.</summary>
    public bool NeedsAttention =>
        Status is FileStatus.Failed or FileStatus.NeedsExcel || NeedsExcel || Residuals.Count > 0;
}
