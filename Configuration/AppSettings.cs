namespace SaveTrigger.Configuration;

/// <summary>
/// Strongly-typed binding for the "AppSettings" section in appsettings.json.
/// All values are read once at startup; live reload is not supported in v1.
/// </summary>
public sealed class AppSettings
{
    /// <summary>Root directories to monitor (recursively).</summary>
    public List<string> MonitoredFolders { get; init; } = [];

    /// <summary>
    /// Zero-based index of the monitor to move Explorer windows to.
    /// Falls back to last available monitor if the index is out of range.
    /// </summary>
    public int TargetMonitorIndex { get; init; } = 0;

    /// <summary>Milliseconds to wait after the last file event in a directory before processing.</summary>
    public int DebounceMs { get; init; } = 1500;

    /// <summary>Maximum number of Explorer windows this app will manage simultaneously.</summary>
    public int MaxManagedWindows { get; init; } = 5;

    /// <summary>Milliseconds between file-size polls during stabilization.</summary>
    public int FileStabilizationPollMs { get; init; } = 500;

    /// <summary>Maximum stabilization attempts before giving up on a file.</summary>
    public int FileStabilizationRetries { get; init; } = 5;

    /// <summary>Seconds of recent foreground activity to consider for local-origin correlation.</summary>
    public int ActivityWindowSeconds { get; init; } = 10;

    /// <summary>File extensions that are always ignored (in-progress downloads, temp writes).</summary>
    public List<string> TempFileExtensions { get; init; } = [];

    /// <summary>Filename prefixes that are always ignored (Office lock files etc.).</summary>
    public List<string> IgnorePrefixes { get; init; } = [];
}
