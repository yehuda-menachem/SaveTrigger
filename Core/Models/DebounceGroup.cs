namespace SaveTrigger.Core.Models;

/// <summary>
/// Accumulates all filesystem events for a single directory within a debounce window.
/// When the window expires (no new events for DebounceMs milliseconds), the group is
/// flushed to the stabilization + correlation stages.
/// </summary>
public sealed class DebounceGroup
{
    public required string DirectoryPath { get; init; }

    /// <summary>All events collected in this window (can include creates and renames).</summary>
    public List<FileEvent> Events { get; } = [];

    public DateTime WindowStart { get; init; }
    public DateTime LastEventAt { get; set; }

    /// <summary>
    /// The event we will use as representative for stabilization and correlation.
    /// Updated to prefer non-temp files as events arrive.
    /// </summary>
    public FileEvent? PrimaryEvent { get; set; }
}
