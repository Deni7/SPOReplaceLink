namespace SPOReplaceLink.State;

public static class RunModes
{
    public const string Inventory = "inventory";
    public const string Scan = "scan";
    public const string Apply = "apply";
    public const string Verify = "verify";

    public static readonly string[] All = [Inventory, Scan, Apply, Verify];

    public static bool Writes(string mode) => mode == Apply;
}

public static class RunStatuses
{
    public const string Created = "created";
    public const string Running = "running";
    public const string Paused = "paused";
    /// <summary>Stopped on request (STA-4); can be continued.</summary>
    public const string Stopped = "stopped";
    /// <summary>Stopped by the breaker, mass auth failures or SPO blocking; can be continued after the cause is fixed.</summary>
    public const string Aborted = "aborted";
    public const string Completed = "completed";
    public const string CompletedWithErrors = "completed_with_errors";

    public static bool IsFinished(string s) => s is Completed or CompletedWithErrors;

    public static bool CanContinue(string s) => s is Created or Stopped or Aborted or Running or Paused;
}

/// <summary>STA-2. Non-final: pending, in_progress, deferred, waiting_path. Everything else is final (RTY-7).</summary>
public static class FileStatuses
{
    public const string WaitingPath = "waiting_path";
    public const string Pending = "pending";
    public const string InProgress = "in_progress";
    public const string Deferred = "deferred";

    public const string Done = "done";
    public const string Unchanged = "unchanged";
    public const string Planned = "planned";
    public const string Verified = "verified";
    /// <summary>verify found links that a run would still change.</summary>
    public const string ChangesPending = "changes_pending";
    public const string NeedsExcel = "needs_excel";
    public const string Ignored = "ignored";
    public const string SkippedFiltered = "skipped:filtered";
    public const string SkippedTooLarge = "skipped:too_large";
    public const string SkippedEncrypted = "skipped:encrypted";
    public const string SkippedDeleted = "skipped:deleted";

    /// <summary>Not a file status: a filter value for all files that need attention (<see cref="StateStore.AttentionSql"/>).</summary>
    public const string Attention = "attention";

    public static string Failed(string errorClass) => "failed:" + errorClass;

    public static readonly string[] NonFinal = [WaitingPath, Pending, InProgress, Deferred];

    public static bool IsFinal(string s) => !NonFinal.Contains(s);
}

/// <summary>A file row as the engine works with it.</summary>
public sealed record FileRow
{
    public long Id { get; init; }
    public long RunId { get; init; }
    public int TargetIndex { get; init; }
    public string DriveId { get; init; } = string.Empty;
    public string ItemId { get; init; } = string.Empty;
    public string? ParentId { get; init; }
    public string Name { get; init; } = string.Empty;
    /// <summary>Path relative to the library root, with '/'.</summary>
    public string? Path { get; init; }
    public long Size { get; init; }
    public string? ETag { get; init; }
    public DateTimeOffset? Modified { get; init; }
    public string? ModifiedBy { get; init; }
    public string Status { get; init; } = FileStatuses.Pending;
    public string? StatusCode { get; init; }
    public string? Stage { get; init; }
    public string? ErrorClass { get; init; }
    public string? Error { get; init; }
    public int Attempts { get; init; }
    public DateTimeOffset? LockedSince { get; init; }
    public string? App { get; init; }
    public string? VersionBefore { get; init; }
    public string? ETagAfter { get; init; }
    public int Replacements { get; init; }
    public int Residuals { get; init; }
    public bool NeedsExcel { get; init; }
    public string? Comment { get; init; }
    /// <summary>eTag before an upload that was started (see StateStore.SetWriteIntent).</summary>
    public string? IntentETag { get; init; }
    /// <summary>SHA-256 of the content that upload sent.</summary>
    public string? IntentHash { get; init; }
}

public sealed record RunRow(
    long Id,
    long? ParentId,
    string Mode,
    string Status,
    string? ProfileName,
    string ConfigSnapshot,
    bool Enumerate,
    DateTimeOffset CreatedAt,
    DateTimeOffset? StartedAt,
    DateTimeOffset? FinishedAt,
    string? StopReason);

public sealed record TargetRow(
    long RunId,
    int Index,
    string? DriveId,
    string? DriveUrl,
    string? RootItemId,
    string? RootPath,
    string? NextLink,
    string? DeltaLink,
    bool Done);

public sealed record FolderRow(string DriveId, string ItemId, string? ParentId, string Name, bool IsRoot);

public sealed record AttemptRow(DateTimeOffset At, int Attempt, string Stage, string? ErrorClass, int? HttpStatus, string? RequestId, string Message, string? App);

/// <summary>RTY-4: which files a retry run takes from its parent.</summary>
public sealed class FileFilter
{
    public List<string> Statuses { get; set; } = [];
    public List<string> ErrorClasses { get; set; } = [];
    public List<string> Extensions { get; set; } = [];
    public string? PathMask { get; set; }
    public List<long> FileIds { get; set; } = [];
    public List<string> Paths { get; set; } = [];
    public int? TargetIndex { get; set; }
    /// <summary>Text search in path (GUI-5).</summary>
    public string? Search { get; set; }
}

/// <summary>What the engine writes when a file reaches a final status or goes back to the queue.</summary>
public sealed record FileOutcome
{
    public required string Status { get; init; }
    public string? StatusCode { get; init; }
    public string? ErrorClass { get; init; }
    public string? Error { get; init; }
    public DateTimeOffset? NotBefore { get; init; }
    public DateTimeOffset? LockedSince { get; init; }
    public string? VersionBefore { get; init; }
    public string? ETagAfter { get; init; }
    public string? ETag { get; init; }
    public int Replacements { get; init; }
    public int Residuals { get; init; }
    public bool NeedsExcel { get; init; }
}
