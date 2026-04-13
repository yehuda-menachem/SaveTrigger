using SaveTrigger.Core.Correlation;

namespace SaveTrigger.Core.Models;

/// <summary>
/// Result of the local-origin correlation check (Stage 4 of the pipeline).
/// Every file event produces a result — rejected events are logged for debugging
/// without triggering any Explorer action.
/// </summary>
public sealed record CorrelationResult(
    /// <summary>True if the file was determined to be created by this machine's user.</summary>
    bool Accepted,

    /// <summary>Human-readable explanation for the accept/reject decision (logged at Debug level).</summary>
    string Reason,

    /// <summary>Creation time of the file as reported by the filesystem.</summary>
    DateTime FileCreatedAt,

    /// <summary>Snapshot of foreground window activity at decision time (for diagnostics).</summary>
    IReadOnlyList<ForegroundSnapshot> RecentActivity
);
