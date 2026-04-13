using System.IO;

namespace SaveTrigger.Core.Models;

/// <summary>
/// Represents a single raw filesystem event captured by FileWatcherService.
/// Immutable record — created once and passed through the pipeline unchanged.
/// </summary>
public sealed record FileEvent(
    /// <summary>Full absolute path to the file (e.g. C:\Downloads\report.pdf).</summary>
    string FullPath,

    /// <summary>Directory containing the file — the folder we will open in Explorer.</summary>
    string DirectoryPath,

    /// <summary>Whether this was a create or rename event.</summary>
    WatcherChangeTypes ChangeType,

    /// <summary>UTC timestamp when our watcher callback fired.</summary>
    DateTime DetectedAt,

    /// <summary>The root folder this watcher was configured for (for diagnostics).</summary>
    string WatcherRoot
);
