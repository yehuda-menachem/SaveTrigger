namespace SaveTrigger.Tray;

/// <summary>
/// Shared mutable state between TrayIconService and the rest of the pipeline.
/// Intentionally kept minimal — only the pause flag is needed in v1.
/// </summary>
public sealed class AppState
{
    /// <summary>
    /// When true, FileWatcherService discards incoming events without enqueuing them.
    /// Written by TrayIconService (STA thread), read by FileWatcherService (thread pool).
    /// Marked volatile to prevent CPU/compiler reordering across threads.
    /// </summary>
    public volatile bool IsPaused = false;
}
