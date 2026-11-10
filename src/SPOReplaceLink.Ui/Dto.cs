using SPOReplaceLink.Engine;
using SPOReplaceLink.State;

namespace SPOReplaceLink.Ui;

public sealed record RunInfo(long Id, long? ParentId, string Mode, string Status, string? ProfileName, DateTimeOffset CreatedAt,
    DateTimeOffset? StartedAt, DateTimeOffset? FinishedAt, string? StopReason)
{
    public static RunInfo From(RunRow r) => new(r.Id, r.ParentId, r.Mode, r.Status, r.ProfileName, r.CreatedAt, r.StartedAt, r.FinishedAt, r.StopReason);
}

public sealed record RunStatus(
    RunInfo Run,
    IReadOnlyDictionary<string, long> Counts,
    long Files,
    long Replacements,
    long Residuals,
    long NonFinal,
    long Attention,
    bool Active,
    bool Paused,
    MetricsSnapshot? Live);


